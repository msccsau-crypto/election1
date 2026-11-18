import os
import threading
from contextlib import asynccontextmanager
from typing import Optional

import openpyxl
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import uvicorn

# ==========================================
# CONFIGURATION
# ==========================================
XLSX_FILE = "voter_list.xlsx"
PHOTOS_FOLDER = "voter_photos"
INDEX_HTML_FILE = "index.html"
VIEWS_COLUMN_HEADER = "VIEWS"  # Single specified column in your .xlsx sheet (Row 1 = VIEWS, Row 2 = count)

# Pre-indexed in-memory stores (Read-only after startup = zero locks for 100+ concurrent users)
VOTERS_BY_ID: dict[int, dict] = {}
VOTERS_INDEX: list[tuple[dict, str, str, str, str]] = []
DEFAULT_TOP_100: list[dict] = []

# Session-based view counter synced to the single cell (Row 2 of VIEWS column) in XLSX_FILE
UNIQUE_SESSIONS: set[str] = set()
TOTAL_VIEWS_COUNT: int = 0
VIEWS_COL_INDEX: int = 10
VIEWS_LOCK = threading.Lock()
SYNC_TIMER: Optional[threading.Timer] = None


class VisitPayload(BaseModel):
    device_id: str


class CachedStaticFiles(StaticFiles):
    """Adds 24-hour browser caching to voter photos so Render's free CPU isn't overloaded."""
    async def get_response(self, path: str, scope):
        response = await super().get_response(path, scope)
        if response.status_code == 200:
            response.headers["Cache-Control"] = "public, max-age=86400, immutable"
        return response


def sync_views_to_xlsx():
    """Background worker: updates ONLY Row 2 of the single VIEWS column in the .xlsx sheet."""
    if not os.path.exists(XLSX_FILE):
        return
    with VIEWS_LOCK:
        count_snapshot = TOTAL_VIEWS_COUNT
        col_idx = VIEWS_COL_INDEX

    try:
        wb = openpyxl.load_workbook(XLSX_FILE)
        ws = wb.active
        ws.cell(row=1, column=col_idx, value=VIEWS_COLUMN_HEADER)
        ws.cell(row=2, column=col_idx, value=int(count_snapshot))

        tmp_xlsx = f"{XLSX_FILE}.tmp.xlsx"
        wb.save(tmp_xlsx)
        wb.close()
        os.replace(tmp_xlsx, XLSX_FILE)
    except PermissionError:
        # Skips silently if the .xlsx is currently open in Excel on your local PC
        pass
    except Exception as e:
        print(f"[XLSX Views Sync Warning]: {e}")


def schedule_views_xlsx_sync():
    """Debounces .xlsx writes so rapid visits write to disk at most once every 3 seconds."""
    global SYNC_TIMER
    if SYNC_TIMER is not None:
        SYNC_TIMER.cancel()
    SYNC_TIMER = threading.Timer(3.0, sync_views_to_xlsx)
    SYNC_TIMER.daemon = True
    SYNC_TIMER.start()


def load_xlsx_into_ram():
    """Loads the .xlsx voter sheet AND reads the initial view count from Row 2 of the VIEWS column."""
    global VOTERS_BY_ID, VOTERS_INDEX, DEFAULT_TOP_100, TOTAL_VIEWS_COUNT, VIEWS_COL_INDEX

    if not os.path.exists(XLSX_FILE):
        raise FileNotFoundError(f"Excel file '{XLSX_FILE}' not found in {os.getcwd()}")

    wb = openpyxl.load_workbook(XLSX_FILE, read_only=True, data_only=True)
    ws = wb.active

    rows_iter = ws.iter_rows(values_only=True)
    raw_headers = next(rows_iter, None)
    if not raw_headers:
        wb.close()
        return

    headers = [str(h or "").strip().upper() for h in raw_headers]

    def find_col(keywords: list[str], fallback: int) -> int:
        for idx, h in enumerate(headers):
            if any(k in h for k in keywords):
                return idx
        return fallback

    col_gpov = find_col(["G.P.O", "GPOV", "LINE", "UNIQUE"], 0)
    col_name = find_col(["VOTER NAME", "NAME"], 1)
    col_rel = find_col(["FATHER", "HUSBAND", "RELATION"], 2)
    col_gender = find_col(["GENDER", "SEX"], 3)
    col_hno = find_col(["H NO", "H.NO", "HOUSE"], 5)
    col_epic = find_col(["VOTER ID", "EPIC"], 6)

    # Locate the single VIEWS column (or assign the next empty column after existing headers)
    if VIEWS_COLUMN_HEADER in headers:
        col_views_zero = headers.index(VIEWS_COLUMN_HEADER)
        VIEWS_COL_INDEX = col_views_zero + 1
    else:
        col_views_zero = -1
        VIEWS_COL_INDEX = len(headers) + 1

    temp_map: dict[int, dict] = {}
    first_data_row = True

    for row in rows_iter:
        if not row:
            continue

        # Read the single-row view count from Row 2 of the VIEWS column
        if first_data_row:
            first_data_row = False
            if col_views_zero != -1 and col_views_zero < len(row) and row[col_views_zero] is not None:
                try:
                    TOTAL_VIEWS_COUNT = int(float(str(row[col_views_zero]).strip()))
                except ValueError:
                    TOTAL_VIEWS_COUNT = 0

        if row[col_gpov] is None:
            continue
        try:
            gpov_no = int(str(row[col_gpov]).strip())
        except ValueError:
            continue

        voter_name = str(row[col_name] or "").strip() if col_name < len(row) else ""
        relation_name = str(row[col_rel] or "").strip() if col_rel < len(row) else ""
        gender = str(row[col_gender] or "").strip().upper() if col_gender < len(row) else ""

        raw_hno = row[col_hno] if col_hno < len(row) and row[col_hno] is not None else ""
        house_no = str(int(raw_hno)) if isinstance(raw_hno, float) and raw_hno.is_integer() else str(raw_hno).strip()

        voter_id = str(row[col_epic] or "").strip() if col_epic < len(row) else ""

        temp_map[gpov_no] = {
            "gpov_no": gpov_no,
            "gender": gender,
            "voter_name": voter_name,
            "relation_name": relation_name,
            "house_no": house_no,
            "voter_id": voter_id,
        }

    wb.close()

    # If the VIEWS column didn't exist yet in the .xlsx, create it with 0 in Row 2
    if col_views_zero == -1:
        sync_views_to_xlsx()

    sorted_ids = sorted(temp_map.keys())
    VOTERS_BY_ID = temp_map
    VOTERS_INDEX = [
        (
            temp_map[gid],
            temp_map[gid]["voter_name"].lower(),
            temp_map[gid]["relation_name"].lower(),
            temp_map[gid]["house_no"].lower(),
            temp_map[gid]["voter_id"].lower(),
        )
        for gid in sorted_ids
    ]
    DEFAULT_TOP_100 = [item[0] for item in VOTERS_INDEX[:100]]
    print(f"Loaded {len(VOTERS_INDEX)} voters and {TOTAL_VIEWS_COUNT} initial views from '{XLSX_FILE}'.")


@asynccontextmanager
async def lifespan(app: FastAPI):
    os.makedirs(PHOTOS_FOLDER, exist_ok=True)
    load_xlsx_into_ram()
    yield


app = FastAPI(
    title="Patharlapalli Voters Check List API",
    docs_url=None,
    redoc_url=None,
    lifespan=lifespan,
)

app.add_middleware(GZipMiddleware, minimum_size=500)

if os.path.exists(PHOTOS_FOLDER):
    app.mount("/voter_photos", CachedStaticFiles(directory=PHOTOS_FOLDER), name="voter_photos")


@app.get("/")
def serve_home():
    if not os.path.exists(INDEX_HTML_FILE):
        raise HTTPException(status_code=404, detail="index.html not found")
    return FileResponse(INDEX_HTML_FILE)


@app.get("/health")
def health_check():
    return {
        "status": "ok",
        "voters_loaded": len(VOTERS_INDEX),
        "total_views": TOTAL_VIEWS_COUNT,
    }


@app.post("/api/views/register")
def register_session_view(payload: VisitPayload):
    """
    Increments TOTAL_VIEWS_COUNT by 1 per new browser session and persists
    it into the single cell (Row 2 of the VIEWS column) in the .xlsx file.
    """
    global TOTAL_VIEWS_COUNT
    session_id = (payload.device_id or "").strip()[:64]
    if not session_id:
        return {"views": TOTAL_VIEWS_COUNT}

    is_new = False
    with VIEWS_LOCK:
        if session_id not in UNIQUE_SESSIONS:
            UNIQUE_SESSIONS.add(session_id)
            TOTAL_VIEWS_COUNT += 1
            is_new = True
        current_total = TOTAL_VIEWS_COUNT

    if is_new:
        schedule_views_xlsx_sync()

    return {"views": current_total}


@app.get("/api/voters/search")
def search_voters(
    gpov_no: Optional[int] = Query(None),
    gender: Optional[str] = Query(None),
    voter_name: Optional[str] = Query(None),
    relation_name: Optional[str] = Query(None),
    house_no: Optional[str] = Query(None),
    voter_id: Optional[str] = Query(None),
    default_limit: Optional[int] = Query(None),
):
    q_gender = gender.strip().upper() if gender else ""
    q_name = voter_name.strip().lower() if voter_name else ""
    q_rel = relation_name.strip().lower() if relation_name else ""
    q_hno = house_no.strip().lower() if house_no else ""
    q_epic = voter_id.strip().lower() if voter_id else ""

    has_other_filters = bool(q_gender or q_name or q_rel or q_hno or q_epic)

    if gpov_no is None and not has_other_filters:
        if default_limit == 100:
            return DEFAULT_TOP_100
        return [item[0] for item in VOTERS_INDEX]

    if gpov_no is not None and not has_other_filters:
        rec = VOTERS_BY_ID.get(gpov_no)
        return [rec] if rec else []

    results = []
    for public_rec, s_name, s_rel, s_hno, s_epic in VOTERS_INDEX:
        if gpov_no is not None and public_rec["gpov_no"] != gpov_no:
            continue
        if q_gender and public_rec["gender"] != q_gender:
            continue
        if q_name and q_name not in s_name:
            continue
        if q_rel and q_rel not in s_rel:
            continue
        if q_hno and q_hno not in s_hno:
            continue
        if q_epic and q_epic not in s_epic:
            continue
        results.append(public_rec)

    return results


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port, workers=1)