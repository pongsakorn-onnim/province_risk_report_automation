from pathlib import Path


def _month_codes(yyyymm: str) -> list[str]:
    year = int(yyyymm[:4])
    start_month = int(yyyymm[4:6])
    codes = []
    for i in range(6):
        m = (start_month - 1 + i) % 12 + 1
        y = year + (start_month - 1 + i) // 12
        codes.append(f"{y}{m:02d}")
    return codes


def get_image_paths(data_dir: Path, region: str, province: str, yyyymm: str) -> dict:
    """Return all image file paths needed to fill the report template."""
    maps = data_dir / "map"
    region_maps = maps / region
    prov_maps = region_maps / province
    gr = data_dir / "graph_region"
    prov_graphs = gr / region
    months = _month_codes(yyyymm)

    return {
        # Slide 4 — national map + bar charts by region
        "nat_all_map":       maps / "all6month.png",
        "nat_flood_graph":   gr / "Flood_รายภาค.png",
        "nat_drought_graph": gr / "Drought_รายภาค.png",
        "nat_both_graph":    gr / "FloodDrought_รายภาค.png",

        # Slide 5 — region map + bar charts by province within region
        "reg_all_map":       region_maps / f"all6month_{region}.png",
        "reg_flood_graph":   gr / f"Flood_{region}.png",
        "reg_drought_graph": gr / f"Drought_{region}.png",
        "reg_both_graph":    gr / f"FloodDrought_{region}.png",

        # Slide 6 — province map + bar charts by amphoe within province
        "prov_all_map":      prov_maps / f"all6month_{province}.png",
        "prov_flood_graph":  prov_graphs / f"Flood_{province}.png",
        "prov_drought_graph":prov_graphs / f"Drought_{province}.png",
        "prov_both_graph":   prov_graphs / f"FloodDrought_{province}.png",

        # Slides 8-13 — province monthly maps (6 pictures, one per month)
        "prov_monthly_maps": [
            prov_maps / f"fcst_union_{mc}_{province}.png" for mc in months
        ],


        # Slide 17 — combined national maps (reuse national 6m maps)
        "comb_flood_map":   maps / "flood6month.png",
        "comb_drought_map": maps / "drought6month.png",
        "comb_both_map":    maps / "flooddrought6month.png",
    }


def get_excel_paths(data_dir: Path, region: str, province: str, yyyymm: str) -> dict:
    province_file = f"{yyyymm}summary_flood_drought_{province}.xlsx"
    province_in_region = data_dir / "excel" / region / province_file
    province_flat     = data_dir / "excel" / province_file
    return {
        "national": data_dir / "excel" / f"{yyyymm}summary_flood_drought.xlsx",
        "province": province_in_region if province_in_region.exists() else province_flat,
    }
