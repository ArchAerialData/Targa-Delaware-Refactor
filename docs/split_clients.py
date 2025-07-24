# pip install docxtpl piexif==1.1.3 python-docx jsonschema pillow
import json
from datetime import datetime
from pathlib import Path
import shutil


def main():
    base = Path(__file__).parent
    src = base / 'clients.json'
    if not src.exists():
        print('clients.json not found')
        return
    with src.open('r') as f:
        data = json.load(f)

    clients_dir = base / 'Clients'
    for client, cfg in data.items():
        cfg_dir = clients_dir / client
        cfg_dir.mkdir(parents=True, exist_ok=True)
        cfg_file = cfg_dir / 'config.json'
        if cfg_file.exists():
            try:
                with cfg_file.open('r') as cf:
                    existing = json.load(cf)
            except Exception:
                existing = None
            if existing == cfg:
                pass
            else:
                with cfg_file.open('w') as cf:
                    json.dump(cfg, cf, indent=2)
        else:
            with cfg_file.open('w') as cf:
                json.dump(cfg, cf, indent=2)

    docs_dir = base / 'docs'
    docs_dir.mkdir(exist_ok=True)
    date_str = datetime.now().strftime('%Y-%m-%d')
    backup = docs_dir / f'clients_master_backup_{date_str}.json'
    if not backup.exists():
        shutil.copy(src, backup)


if __name__ == '__main__':
    main()
