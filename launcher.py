import os
import sys
import logging
from pathlib import Path

import yaml
import questionary

# When running as a built .exe, __file__ doesn't exist — use exe location instead
if getattr(sys, 'frozen', False):
    ROOT = Path(sys.executable).parent
else:
    ROOT = Path(__file__).parent

# Set working directory to project root so all relative paths inside ppt_builder work
os.chdir(ROOT)

ALL_PROVINCES = sorted([
    "กระบี่", "กรุงเทพมหานคร", "กาญจนบุรี", "กาฬสินธุ์", "กำแพงเพชร",
    "ขอนแก่น", "จันทบุรี", "ฉะเชิงเทรา", "ชลบุรี", "ชัยนาท",
    "ชัยภูมิ", "ชุมพร", "เชียงราย", "เชียงใหม่", "ตรัง",
    "ตราด", "ตาก", "นครนายก", "นครปฐม", "นครพนม",
    "นครราชสีมา", "นครศรีธรรมราช", "นครสวรรค์", "นนทบุรี", "นราธิวาส",
    "น่าน", "บึงกาฬ", "บุรีรัมย์", "ปทุมธานี", "ประจวบคีรีขันธ์",
    "ปราจีนบุรี", "ปัตตานี", "พระนครศรีอยุธยา", "พะเยา", "พังงา",
    "พัทลุง", "พิจิตร", "พิษณุโลก", "เพชรบุรี", "เพชรบูรณ์",
    "แพร่", "ภูเก็ต", "มหาสารคาม", "มุกดาหาร", "แม่ฮ่องสอน",
    "ยโสธร", "ยะลา", "ร้อยเอ็ด", "ระนอง", "ระยอง",
    "ราชบุรี", "ลพบุรี", "ลำปาง", "ลำพูน", "เลย",
    "ศรีสะเกษ", "สกลนคร", "สงขลา", "สตูล", "สมุทรปราการ",
    "สมุทรสงคราม", "สมุทรสาคร", "สระแก้ว", "สระบุรี", "สิงห์บุรี",
    "สุโขทัย", "สุพรรณบุรี", "สุราษฎร์ธานี", "สุรินทร์", "หนองคาย",
    "หนองบัวลำภู", "อ่างทอง", "อำนาจเจริญ", "อุดรธานี", "อุตรดิตถ์",
    "อุทัยธานี", "อุบลราชธานี",
])

REGIONS = [
    "เหนือ",
    "ตะวันออกเฉียงเหนือ",
    "กลาง",
    "ตะวันออก",
    "ใต้",
    "ลุ่มน้ำเจ้าพระยา",
]


def load_config() -> dict:
    with open(ROOT / "config.yaml", encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_data_base_dir(cfg: dict) -> Path:
    """
    Primary: relative path from project folder to data folder (works for all PCs
    once the project is in OneDrive under HydroDataSci\Project\).
    Fallback: data_base_dir from config.yaml (used during development).
    """
    relative = ROOT.parent.parent / "Data" / "Risk_Area" / "Risk_Forecast" / "summary"
    if relative.exists():
        return relative
    fallback = cfg.get("data_base_dir", "")
    if fallback:
        return Path(fallback)
    return relative


def scan_province_region(data_dir: Path, yyyymm: str) -> dict:
    """Return {province: region} by scanning excel/ subdirectories."""
    excel_dir = data_dir / "excel"
    mapping = {}
    if not excel_dir.exists():
        return mapping
    for region_dir in excel_dir.iterdir():
        if not region_dir.is_dir():
            continue
        for f in region_dir.glob(f"{yyyymm}summary_flood_drought_*.xlsx"):
            province = f.stem.replace(f"{yyyymm}summary_flood_drought_", "")
            mapping[province] = region_dir.name
    for f in excel_dir.glob(f"{yyyymm}summary_flood_drought_*.xlsx"):
        province = f.stem.replace(f"{yyyymm}summary_flood_drought_", "")
        if province not in mapping:
            mapping[province] = None
    return mapping


def main():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    cfg = load_config()
    data_base_dir = get_data_base_dir(cfg)

    print()
    print("=" * 46)
    print("   Province Risk Report Generator")
    print("=" * 46)
    print()

    # Forecast month
    yyyymm = questionary.text(
        "Forecast month (YYYYMM, e.g. 202606):",
        validate=lambda v: True if (len(v) == 6 and v.isdigit()) else "Enter 6 digits, e.g. 202606",
    ).ask()
    if yyyymm is None:
        sys.exit(0)

    data_dir = data_base_dir / f"{yyyymm}_summary"
    if not data_dir.exists():
        print(f"\n  [!] Data directory not found:\n      {data_dir}")
        input("\n  Press Enter to close...")
        sys.exit(1)

    available = scan_province_region(data_dir, yyyymm)

    # Province
    province = questionary.autocomplete(
        "Province (type to filter):",
        choices=ALL_PROVINCES,
        match_middle=True,
        validate=lambda v: True if v in ALL_PROVINCES else "Select a province from the list",
    ).ask()
    if province is None:
        sys.exit(0)

    # Region: auto-detect from data dir, or ask
    region = available.get(province)
    if province not in available:
        print(f"\n  [!] No data found for {province} in {yyyymm}.")
        print("      Slides with missing data will show placeholder images.\n")
        region = questionary.select("Region:", choices=REGIONS).ask()
        if region is None:
            sys.exit(0)
    elif region is None:
        region = questionary.select(
            "Region (could not auto-detect — flat data structure):",
            choices=REGIONS,
        ).ask()
        if region is None:
            sys.exit(0)

    # Confirm
    print()
    print(f"  Forecast month : {yyyymm}")
    print(f"  Province       : {province}")
    print(f"  Region         : {region}")
    print()

    try:
        input("  Press \033[1;32mEnter\033[0m to generate report (\033[33mCtrl+C\033[0m to cancel)  ")
    except KeyboardInterrupt:
        print("\n\n  Cancelled.")
        sys.exit(0)

    # Generate
    print()
    try:
        from src.ppt_builder import build_report
        year = int(yyyymm[:4])
        start_month = int(yyyymm[4:6])
        out = build_report(
            config=cfg,
            province=province,
            region=region,
            data_dir=data_dir,
            yyyymm=yyyymm,
            year=year,
            start_month=start_month,
            output_dir=ROOT / "output",
        )
        print()
        print("  Done! Saved to:")
        print(f"  {out}")
    except Exception as e:
        print(f"\n  [!] Report generation failed: {e}")

    input("\n  Press Enter to close...")


if __name__ == "__main__":
    main()
