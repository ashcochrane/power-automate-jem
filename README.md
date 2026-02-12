# Jem — Eden Park Mobility Parking Pass Automation

Automates the generation and emailing of personalised mobility parking passes for Eden Park events.

## How it works

1. Ballot entries come in as a CSV from the Eden Park website
2. The operator pastes the data into Excel and removes duplicates
3. Power Automate generates a personalised PDF pass for each row and emails it to the recipient

## What's in this repo

| Folder | Contents |
|---|---|
| `templates/` | Word template (with content controls), email body text |
| `solution/` | Importable Power Automate flow package (.zip), standalone flow definition (.json), and import instructions |
| `docs/` | Setup guides for IT, the end user, the developer, and a full manual build guide |
| `samples/` | Test data (fake CSV rows) |

### Key files

| File | Who it's for |
|---|---|
| [`solution/MobilityParkingPass.zip`](solution/MobilityParkingPass.zip) | IT — importable flow package |
| [`solution/IMPORT-INSTRUCTIONS.md`](solution/IMPORT-INSTRUCTIONS.md) | IT — step-by-step import guide |
| [`docs/it-setup-guide.md`](docs/it-setup-guide.md) | IT — full setup: folders, templates, import, configure, test |
| [`docs/end-user-guide.md`](docs/end-user-guide.md) | Operator — paste data, dedup, run, check results |
| [`docs/flow-build-guide.md`](docs/flow-build-guide.md) | Developer — rebuild the flow from scratch if needed |
| [`docs/manual-build-guide.md`](docs/manual-build-guide.md) | Developer — complete step-by-step build guide with screenshots-level detail |

## Setup

See [`docs/it-setup-guide.md`](docs/it-setup-guide.md) for full import and configuration instructions.

## Usage

See [`docs/end-user-guide.md`](docs/end-user-guide.md) for the end user workflow.
