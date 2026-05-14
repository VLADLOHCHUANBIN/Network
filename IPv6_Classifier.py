import os
import glob
import pandas as pd
import ipaddress
import re
from collections import defaultdict

# --- CONFIGURATION ---
FAMILY_LOGS_FOLDER = r"D:\2Work Task\IPv6 Collection\Family Log\**\*.log"
TRAFFIC_LOGS_FOLDER = r"D:\2Work Task\IPv6 Collection\Log\**\*.log"
OUTPUT_EXCEL_FILE = r"D:\2Work Task\IPv6 Collection\IPv6_Traffic_Family_Counts2.xlsx"

# NEW: How large should we group the unknown traffic?
# 48 is standard for external routing. 64 is standard for single LANs.
UNKNOWN_GROUPING_PREFIX = 48
# ---------------------

family_extractor = re.compile(r"address\s+\d+\s+([a-fA-F0-9:]+)\s+(\d+)\s+description\s+(\S+)", re.IGNORECASE)
ip_extractor = re.compile(r"([a-fA-F0-9:]+)\.\d+\s+-->")


def process_logs(family_folder_path, traffic_folder_path):
    print(f"Locating Family files in: {family_folder_path}")
    family_file_list = glob.glob(family_folder_path, recursive=True)

    print(f"Locating Traffic log files in: {traffic_folder_path}")
    traffic_file_list = glob.glob(traffic_folder_path, recursive=True)

    if not family_file_list or not traffic_file_list:
        print("Error: Could not find log files in one or both directories!")
        return None

    # --- PASS 1: Build the Family Map ---
    print("\n--- PASS 1: Extracting Switch Families ---")
    fast_map = defaultdict(dict)
    family_lines_found = 0

    for filepath in family_file_list:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                match = family_extractor.search(line)
                if match:
                    ip = match.group(1)
                    prefix = int(match.group(2))
                    desc = match.group(3)

                    try:
                        network = ipaddress.IPv6Network(f"{ip}/{prefix}", strict=False)
                        if prefix <= 44:
                            for sub in network.subnets(new_prefix=44):
                                fast_map[44][sub] = desc
                        else:
                            fast_map[prefix][network] = desc
                        family_lines_found += 1
                    except ValueError:
                        pass

    print(f"Found {family_lines_found} family routing lines.")

    # --- PASS 2: Match and Count Traffic ---
    print("\n--- PASS 2: Counting Active Traffic Sessions ---")

    sorted_prefixes = sorted(fast_map.keys(), reverse=True)
    counts = defaultdict(lambda: defaultdict(int))
    traffic_lines_processed = 0
    matched_ips = 0

    for filepath in traffic_file_list:
        print(f"  -> Scanning: {os.path.basename(filepath)}")
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as file:
            for line in file:
                if "address" in line and "description" in line:
                    continue

                match = ip_extractor.search(line)
                if match:
                    raw_ip_str = match.group(1)

                    try:
                        is_matched = False

                        for pref in sorted_prefixes:
                            masked_net = ipaddress.IPv6Network(f"{raw_ip_str}/{pref}", strict=False)
                            if masked_net in fast_map[pref]:
                                desc = fast_map[pref][masked_net]
                                counts[desc][str(masked_net)] += 1
                                is_matched = True
                                matched_ips += 1
                                break

                                # NEW CATEGORIZATION LOGIC
                        if not is_matched:
                            # Group the unmatched IP into a /48 subnet block
                            unknown_net = ipaddress.IPv6Network(f"{raw_ip_str}/{UNKNOWN_GROUPING_PREFIX}", strict=False)
                            counts['UNMATCHED_TRAFFIC'][str(unknown_net)] += 1

                        traffic_lines_processed += 1

                        if traffic_lines_processed % 500000 == 0:
                            print(f"    ...processed {traffic_lines_processed:,} traffic IPs so far.")

                    except ValueError:
                        pass

    print(f"\nDone! Processed {traffic_lines_processed:,} total traffic sessions.")
    print(f"Successfully matched {matched_ips:,} to a Switch Family.")
    return counts


def export_to_excel(counts, output_filename):
    if not counts:
        return

    print("\nGenerating multi-sheet Excel report...")
    rows = []
    for desc, subnets in counts.items():
        for sub, count in subnets.items():
            rows.append({
                'Switch Family': desc,
                'Subnet Block': sub,
                'Active Sessions (Count)': count
            })

    df = pd.DataFrame(rows)
    df = df.sort_values(by=['Switch Family', 'Active Sessions (Count)'],
                        ascending=[True, False])  # Sort by highest count!

    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        switches = df['Switch Family'].unique()
        for switch in switches:
            switch_data = df[df['Switch Family'] == switch]
            safe_sheet_name = str(switch)[:31]
            switch_data.to_excel(writer, index=False, sheet_name=safe_sheet_name)

            worksheet = writer.sheets[safe_sheet_name]
            for col in worksheet.columns:
                max_len = 0
                column_letter = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_len:
                            max_len = len(str(cell.value))
                    except:
                        pass
                worksheet.column_dimensions[column_letter].width = max_len + 5

    print(f"Success! Output saved to: {output_filename}")


if __name__ == "__main__":
    traffic_counts = process_logs(FAMILY_LOGS_FOLDER, TRAFFIC_LOGS_FOLDER)
    if traffic_counts:
        export_to_excel(traffic_counts, OUTPUT_EXCEL_FILE)
