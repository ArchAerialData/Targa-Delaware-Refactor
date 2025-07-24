The goal of this PR / Codespace session is purely structural:

Split the monolithic clients.json into one config.json per client under Clients/<CLIENT>/.

Patch the existing Python so it loads the same data, runs the same reports, and produces identical output.

Leave room for future per-client customisations (e.g., HAT Intersection Reports) but do not add new report logic now.

Important: functionality and output must remain **bit-for-bit identical** after the migration.
This change is only about file organisation and small path-handling tweaks.

Current repo overview
root/
├── clients.json                # single master config (today)
├── report_generator.py         # main script
├── Clients/
│   ├── HAT/                    # already contains templates & KMZs
│   ├── ENST/
│   └── …
└── docs/

Target layout (after Copilot finishes)
root/
├── split_clients.py            # one-shot helper (idempotent)
├── report_generator.py         # patched to load per-client configs
├── Clients/
│   ├── HAT/
│   │   ├── config.json         # extracted from old master file
│   │   └── Templates - DOCX/
│   ├── ENST/
│   │   └── config.json
│   └── …
└── docs/
    └── clients_master_backup_YYYY-MM-DD.json   # archived original

Target layout (after Copilot finishes)
root/
├── split_clients.py            # one-shot helper (idempotent)
├── report_generator.py         # patched to load per-client configs
├── Clients/
│   ├── HAT/
│   │   ├── config.json         # extracted from old master file
│   │   └── Templates - DOCX/
│   ├── ENST/
│   │   └── config.json
│   └── …
└── docs/
    └── clients_master_backup_YYYY-MM-DD.json   # archived original
🔌 Required Python libraries
For new contributors on a clean machine, every script that imports external packages should begin with this commented one-liner:

python
# pip install docxtpl piexif==1.1.3 python-docx jsonschema pillow
Add that comment near the top of both report_generator.py and the new split_clients.py.
The same line should also appear in `requirements.txt` so contributors can install everything in one step.

⬇ Migration steps Copilot must perform
Copilot: complete all of the following in a single branch / PR.

Generate split_clients.py

Read clients.json.

For each top-level key ("HAT", "ENST", …):

Create Clients/<KEY>/config.json (make parent dirs if missing).

Dump the block unchanged.

Backup the original file to `docs/clients_master_backup_YYYY-MM-DD.json` using today's date.

Make the script idempotent – running it twice should not duplicate work.

Patch report_generator.py

Replace the hard-coded clients.json path with

python
cfg_file = Path(__file__).parent / "Clients" / client_code / "config.json"
Remove any fallback that auto-loads the first .json it finds.

Update the GUI / CLI client picker

List folders under Clients/ that contain a config.json.

Run `split_clients.py` first to generate per-client configs, then delete the obsolete `clients.json` once CI verifies outputs match pre-migration runs.

Add a JSON Schema (schemas/client_config.schema.json) and wire simple validation.

Append a “How to add a new client” section to this README.

Create / update requirements.txt mirroring the pip line above.

Verification checklist
 Unit tests pass.

 Running python split_clients.py twice is a no-op the second time.

 Sample reports match pre-migration output (binary diff).

 GUI lists clients dynamically and runs end-to-end.

After merge
Add a client by copying an existing folder, renaming it, and editing its config.json.

## How to add a new client
1. Copy any existing folder under `Clients/` and rename it to the new client code.
2. Edit `config.json` inside that folder to match the new client's settings.
3. Place any templates or KMZ files in the appropriate subfolders.

Future enhancements will extend the JSON Schema and generator; today’s change is structural only.

Happy migrating! 🚀

## Running the New GUI

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Execute the report generator (you can also drag a folder onto the script on Windows to prefill the folder field):

   ```bash
   python report_generator.py
   ```

   A window will appear centered at 1600x900 pixels. Select a folder of photos,
   choose the client and cover photo, pick up to three pilots, then click
   **Generate Reports**.

The GUI will display an optional logo if an image file named `arch_logo.png` is
present next to `report_generator.py`. Edit `report_generator.py` to point to a
different logo path if desired.

