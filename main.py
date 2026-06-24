import argparse
import sys
import io
import logging
from pathlib import Path
import yaml

# Fix Thai output in Windows terminal
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(
        description="Build drought/flood province forecast report"
    )
    parser.add_argument("--province", required=True,
                        help="Province name in Thai, e.g. น่าน")
    parser.add_argument("--region", required=True,
                        help="Region directory name in Thai, e.g. เหนือ")
    parser.add_argument("--data-dir", required=True,
                        help="Path to the 202606_summary data directory")
    parser.add_argument("--yyyymm", default="202606",
                        help="Year-month of the forecast period, default 202606")
    parser.add_argument("--output-dir", default="output",
                        help="Directory to write the output .pptx, default ./output")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
        stream=sys.stderr,
    )

    with open("config.yaml", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    year = int(args.yyyymm[:4])
    start_month = int(args.yyyymm[4:6])

    from src.ppt_builder import build_report

    out = build_report(
        config=config,
        province=args.province,
        region=args.region,
        data_dir=Path(args.data_dir),
        yyyymm=args.yyyymm,
        year=year,
        start_month=start_month,
        output_dir=Path(args.output_dir),
    )
    print(f"Report saved: {out}")


if __name__ == "__main__":
    main()
