import os
import json
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
XLSX_FILE = "Voter_List.xlsx"
PHOTOS_FOLDER = "voter_photos"
INDEX_HTML_FILE = "index.html"
VIEWS_STATE_FILE = "site_views.json"

# Pre-indexed in-memory stores (Read-only after startup = zero locks for 100+ concurrent users)
VOTERS_BY_ID: dict[int, dict] = {}
VOTERS_INDEX: list[tuple[dict, str, str, str, str]] = []
DEFAULT_TOP_100: list[dict] = []

# Unique device/session view tracker (in RAM + lightweight JSON persistence)
UNIQUE_DEVICES: set[str] = set()
VIEWS_LOCK = threading.Lock()


class VisitPayload(BaseModel):
    device_id: str


class CachedStaticFiles(StaticFiles):
    """Adds 24-hour browser caching to voter photos so Render's free CPU isn't overloaded."""
    async def get_response(self, path: str, scope):
        response = await super().get_response(path, scope)
        if response.status_code == 200:
            response.headers["Cache-Control"] = "public, max-age=86400, immutable"
        return response


def load_views_from_disk():
    """Loads previously recorded unique device IDs if available on disk."""
    global UNIQUE_DEVICES
    if not os.path.exists(VIEWS_STATE_FILE):
        return
    try:
        with open(VIEWS_STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            UNIQUE_DEVICES = {str(x) for x in data.get("devices", []) if x}
    except Exception:
        UNIQUE_DEVICES = set()


def save_views_to_disk(devices_snapshot: list[str]):
    """Atomically persists unique device IDs to local JSON."""
    try:
        tmp_file = f"{VIEWS_STATE_FILE}.tmp"
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump({"total_views": len(devices_snapshot), "devices": devices_snapshot}, f)
        os.replace(tmp_file, VIEWS_STATE_FILE)
    except Exception:
        pass


def load_xlsx_into_ram():
    """Loads the .xlsx voter sheet once into RAM at startup and pre-builds response objects."""
    global VOTERS_BY_ID, VOTERS_INDEX, DEFAULT_TOP_100

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

    temp_map: dict[int, dict] = {}
    for row in rows_iter:
        if not row or row[col_gpov] is None:
            continue
        try:
            gpov_no = int(str(row[col_gpov]).strip())
        except ValueError:
            continue

        voter_name = str(row[col_name] or "").strip() if col_name < len(row) else ""
        relation_name = str(row[col_rel] or "").strip() if col_rel < len(row) else ""
        gender = str(row[col_gender] or "").strip().upper() if col_gender < len(row) else ""
        house_no = str(row[col_hno] or "").strip() if col_hno < len(row) else ""
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
    print(f"Loaded {len(VOTERS_INDEX)} voters from '{XLSX_FILE}' into RAM.")


@asynccontextmanager
async def lifespan(app: FastAPI):
    os.makedirs(PHOTOS_FOLDER, exist_ok=True)
    load_views_from_disk()
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
        "total_views": len(UNIQUE_DEVICES),
    }


@app.post("/api/views/register")
def register_device_view(payload: VisitPayload):
    """
    Counts 1 view per unique browser session. Repeated searches or refreshes
    within the same session return the current total without incrementing again.
    """
    dev_id = (payload.device_id or "").strip()[:64]
    if not dev_id:
        return {"views": len(UNIQUE_DEVICES)}

    is_new = False
    with VIEWS_LOCK:
        if dev_id not in UNIQUE_DEVICES:
            UNIQUE_DEVICES.add(dev_id)
            is_new = True
        total_views = len(UNIQUE_DEVICES)
        snapshot = list(UNIQUE_DEVICES) if is_new else None

    if is_new and snapshot is not None:
        save_views_to_disk(snapshot)

    return {"views": total_views}


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