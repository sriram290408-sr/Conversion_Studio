# Power BI Metadata Converter

A local web interface for converting Power BI PBIX metadata into Excel workbooks and, on Windows, creating an interactive live Power BI connection inside Excel.

## What this app does

This repository contains:

- a FastAPI backend in `app/main.py`
- a static React-based UI in the repository root
- a Windows-only live connect workflow using Excel COM automation
- a file upload and analysis pipeline for PBIX metadata

The UI can run in two modes:

1. **Standard conversion mode** – uploads PBIX metadata and returns conversion previews and downloads.
2. **Live Excel mode** – uploads source files to the local Windows agent, opens Excel, and builds a live Power BI-connected workbook.

## How the website workflow works

### 1. Startup and backend detection

When the page loads, the frontend checks two API contexts:

- `window.apiURL(...)` → standard backend API
- `window.liveApiURL(...)` → local Windows agent API

If the browser is hosted locally (`localhost` or `127.0.0.1`), the UI prefers the local agent and runs in `Local Windows mode`.

The UI periodically polls `/api/health` on both endpoints to determine:

- whether the backend is available
- whether the Windows agent is online
- whether live connect support is ready

### 2. Required input files

The main workflow requires three input files:

- `.pbix` file
- dashboard screenshot (`.png`, `.jpg`, `.jpeg`, or `.webp`)
- TMDL metadata file (`.tmdl`, `.txt`, or `.json`)

These files are uploaded from the left panel in the web UI.

### 3. Standard conversion flow

When the user clicks the primary conversion button, the UI:

- uploads the selected files to the backend via `POST /api/upload`
- triggers metadata analysis
- receives a preview of extracted chunks, tables, visuals, and relationships
- displays a structured preview in the UI
- allows downloading a PDF preview or converted output

### 4. Live Excel connection flow (Windows only)

If a local Windows agent is available, the UI enables the live connect workflow.

The live connect flow is:

1. Upload the same selected source files directly to the local agent endpoint `/upload`
2. Start live connect with `POST /live-connect/start`
3. The backend prepares files and launches Excel
4. The user opens Excel and creates a PivotTable from the Power BI semantic model
5. The UI polls `/live-connect/{session_id}/status` to follow the workflow
6. The user clicks `Continue` after Excel is ready
7. The agent verifies the model match and finalizes the live workbook
8. Success allows downloading the live Excel workbook and report

### 5. Error handling

The web UI shows clear user feedback for:

- missing required files
- failed uploads
- live agent offline
- Excel connection detection failures
- semantic model mismatch
- live connect cancellation

## Installation

1. Open a command prompt in the repository root.
2. Create and activate a Python virtual environment if needed:

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

3. Install dependencies:

```cmd
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Running the app locally

Start the FastAPI backend from the project root:

```cmd
.venv\Scripts\python.exe -m uvicorn app.main:app --host 127.0.0.1 --port 8000
```

Then open the website in a browser at:

```txt
http://127.0.0.1:8000/
```

## Local Windows agent workflow

For live Excel COM support, start the local Windows agent first by running the batch file from the project folder:

```cmd
start-local-agent.bat
```

Keep the terminal open. The frontend will detect the local agent and show live connect controls.

## Main frontend pages and features

- `index.html` – single-page UI shell
- `index.js` – frontend React application and user interaction logic
- `api.js` – API endpoint resolver for cloud vs local agent
- `index.css` – styling and layout

## Important API endpoints

The backend exposes the following user-facing endpoints:

- `GET /api/health` and `GET /health`
  - basic application health status
- `GET /api/system-check` and `GET /system-check`
  - diagnostics for the running agent
- `POST /upload`
  - upload required PBIX metadata, screenshot, and TMDL files
- `POST /api/live-connect/start` and `POST /live-connect/start`
  - instruct the Windows agent to begin live Excel connection
- `GET /live-connect/{session_id}/status`
  - poll current live session state
- `POST /live-connect/{session_id}/continue`
  - tell the agent to continue after Excel setup
- `POST /live-connect/{session_id}/cancel`
  - cancel an active live session
- `GET /live-connect/{session_id}/report` and `/download`
  - download live session report or finished workbook

## Deployment notes for agents

- The UI supports a cloud-hosted standard backend and a local Windows agent.
- On cloud-hosted pages, the app tries to use the cloud API first and falls back to the local Windows agent if available.
- On `localhost`, the UI prioritizes the local Windows agent for live connect.
- The local Windows agent is required for Excel COM interaction. Without it, only standard metadata analysis and conversion are available.

## Usage summary

1. Start the backend locally with `uvicorn`.
2. Open the UI in a browser at `http://127.0.0.1:8000/`.
3. Upload a PBIX file, dashboard screenshot, and TMDL metadata.
4. Click the conversion button.
5. If Windows live connect is available, follow the live connection prompts in Excel.
6. Download the live Excel workbook or preview report after completion.

## Requirements

- Windows 10 or later for live connect functionality
- Python 3.12 or compatible
- Microsoft Excel installed for live connect
- The project contains a `.venv` directory for the local Python environment

## Troubleshooting

- If the page says the Windows agent is offline, start `start-local-agent.bat` and keep the terminal open.
- If the backend fails to import, confirm you run `python -m uvicorn app.main:app` from the repository root.
- If live connect fails during Excel use, ensure Excel can open and connect to the Power BI semantic model.

## Notes

- The app is designed to work as both a local Windows desktop agent and a web front-end.
- For fully interactive live Excel functionality, the Windows agent and Excel are required.
- If you do not have Windows or Excel, the site still supports standard PBIX metadata conversion and preview.

## License

This repository does not include an explicit license file. Add one if you plan to share or distribute the project.
