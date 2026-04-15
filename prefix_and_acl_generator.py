import ipaddress

# Mapping Areas (Telcos) to simple identifiers
AREAS = ["A", "B", "C", "D", "E", "F", "G", "H"]

# Type Data keyed by VPN Instance with the requested IP placeholders
TYPE_DATA = {
    "AA": {
        "ips": ["11.11.11.160/27", "11.11.11.192/27", "11.11.11.224/27"]
    },
    "BB": {
        "ips": ["22.22.22.113/32", "22.22.22.114/32", "22.22.22.115/32", "22.22.22.116/32"]
    },
    "CC": {
        "ips": ["33.33.33.140/29", "33.33.33.148/29", "33.33.33.156/29"]
    }
}

def get_in_ips():
    print("\n--- Enter Remote (IN) IPs ---")
    print("Format: IP Mask (e.g., 115.164.118.16 32)")
    print("Type 'done' to finish.")
    ips = []
    while True:
        entry = input("IP and Mask: ").strip()
        if entry.lower() == 'done': break
        if entry:
            formatted_entry = entry.replace(" ", "/")
            try:
                ipaddress.IPv4Network(formatted_entry, strict=False)
                ips.append(formatted_entry)
            except ValueError:
                print("Invalid format. Use 'IP Mask' (e.g., 1.1.1.0 24)")
    return ips

def main():
    # 1. Selection: Area and VPN Instance
    print("\nSelect Area:")
    for i, a in enumerate(AREAS, 1): print(f"{i}. Area {a}")
    selected_area = AREAS[int(input("Choice #: ")) - 1]

    vpn_list = list(TYPE_DATA.keys())
    print("\nSelect VPN Instance:")
    for i, v in enumerate(vpn_list, 1): print(f"{i}. {v}")
    selected_vpn = vpn_list[int(input("Choice #: ")) - 1]
    
    out_ips = TYPE_DATA[selected_vpn]["ips"]

    # 2. Index Requirements
    print("\n--- Indexing Configuration ---")
    idx_out_start = int(input("Start index for Prefix-OUT: "))
    rule_start = int(input("Start index for ACL Rules: "))

    # 3. Input Remote IPs
    in_ips = get_in_ips()

    # Build VRP Names
    prefix_name_in = f"{selected_area}_{selected_vpn}_IN"
    prefix_name_out = f"{selected_area}_{selected_vpn}_OUT"
    acl_name = f"{selected_area}_{selected_vpn}_EF"

    print("\n" + "=" * 75)
    print(f" FINAL CONFIG: Area {selected_area} | VPN {selected_vpn}")
    print("=" * 75)

    # SECTION 1: IP-PREFIX IN
    print("\n# [SECTION 1: IP-PREFIX IN]")
    curr_idx_in = 10
    for ip in in_ips:
        network, mask = ip.split('/')
        print(f"ip ip-prefix {prefix_name_in} index {curr_idx_in} permit {network} {mask}")
        curr_idx_in += 10

    # SECTION 2: IP-PREFIX OUT
    print("\n# [SECTION 2: IP-PREFIX OUT]")
    curr_idx_out = idx_out_start
    for ip in out_ips:
        network, mask = ip.split('/')
        print(f"ip ip-prefix {prefix_name_out} index {curr_idx_out} permit {network} {mask}")
        curr_idx_out += 10

    # SECTION 3: ADVANCED ACL
    print(f"\n# [SECTION 3: ADVANCED ACL]")
    print(f"acl name {acl_name} advance")
    curr_rule = rule_start
    for in_ip_str in in_ips:
        in_iface = ipaddress.IPv4Interface(in_ip_str)
        for out_ip_str in out_ips:
            out_iface = ipaddress.IPv4Interface(out_ip_str)
            print(f" rule {curr_rule} permit ip source {in_iface.ip} {in_iface.network.hostmask} "
                  f"destination {out_iface.ip} {out_iface.network.hostmask}")
            curr_rule += 5

    print("\n" + "=" * 75)

if __name__ == "__main__":
    main()
