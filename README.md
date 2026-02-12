# Jem — Eden Park Mobility Parking Pass Automation

Automates the generation and emailing of personalised mobility parking passes for Eden Park events.

## How it works

1. Ballot entries come in as a CSV from the Eden Park website
2. The operator pastes the data into Excel and removes duplicates
3. Power Automate generates a personalised PDF pass for each row and emails it to the recipient

## What's in this repo

| Folder | Contents |
|---|---|
| `templates/` | Word template (with content controls), Excel starter file setup guide, email body text |
| `solution/` | Exported Power Automate solution package (.zip) |
| `docs/` | Setup guides for IT, the end user, and the developer |
| `samples/` | Test data and example output |

## Setup

See [`docs/it-setup-guide.md`](docs/it-setup-guide.md) for full import and configuration instructions.

## Usage

See [`docs/end-user-guide.md`](docs/end-user-guide.md) for the end user workflow.
