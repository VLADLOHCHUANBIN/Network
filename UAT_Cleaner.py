import os

# ==========================================
# 🛑 CONFIGURATION ZONE 🛑
# ==========================================
SITES = ["A", "B", "C"]


# ==========================================

def surgical_uat_cleanup():
    base_dir = os.getcwd()

    for site in SITES:
        site_path = os.path.join(base_dir, site)
        if not os.path.exists(site_path):
            continue

        print(f"\n--- Cleaning Site: {site} ---")

        for router_folder in os.listdir(site_path):
            router_path = os.path.join(site_path, router_folder)
            if not os.path.isdir(router_path):
                continue

            # Identify if this folder belongs to a Gateway or a Switch
            # (Matches: DCGW, SW, CTOR, MTOR, UTOR, EOR)
            is_gateway = "GW" in router_folder.upper()
            is_switch = any(x in router_folder.upper() for x in ["SW", "TOR", "EOR"])

            print(f"[*] Folder: {router_folder} (Detected as: {'Gateway' if is_gateway else 'Switch'})")

            for filename in os.listdir(router_path):
                # 1. ALWAYS skip HAT files and temp files
                if "HAT" in filename.upper() or filename.startswith("~$"):
                    continue

                # 2. ONLY look at UAT files
                if "UAT" in filename.upper():
                    file_path = os.path.join(router_path, filename)
                    fn_upper = filename.upper()

                    # --- THE REMOVAL LOGIC ---

                    # Case A: It's a Switch folder, but we found a Gateway (NE8000) UAT
                    if is_switch and "NE8000" in fn_upper:
                        print(f"    [DELETE] Removing unrelated Gateway UAT: {filename}")
                        os.remove(file_path)

                    # Case B: It's a Gateway folder, but we found a Switch (CE6800) UAT
                    elif is_gateway and "CE6800" in fn_upper:
                        print(f"    [DELETE] Removing unrelated Switch UAT: {filename}")
                        os.remove(file_path)

    print("\n✅ Cleanup Complete! Wrong UAT versions removed, HATs preserved.")


if __name__ == "__main__":
    surgical_uat_cleanup()
