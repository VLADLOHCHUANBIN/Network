import ipaddress

# Fixed Data Structures
TELCOS = ["CELCOM", "DIGI", "REDTONE", "TIME", "TM", "UM", "WEBE", "YTL"]
SERVICES = {
    "RTP": [
        "202.122.147.160/27",
        "202.122.147.192/27",
        "202.122.147.224/27"
    ],
    "SIGTRAN": [
        "202.122.147.113/32",
        "202.122.147.114/32",
        "202.122.147.115/32",
        "202.122.147.116/32"
    ],
    "SIP": [
        "202.122.147.140/29",
        "202.122.147.148/29",
        "202.122.147.156/29"
    ]
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
    # 1. Selection
    print("\nSelect Telco:")
    for i, t in enumerate(TELCOS, 1): print(f"{i}. {t}")
    selected_telco = TELCOS[int(input("Choice #: ")) - 1]

    service_list = list(SERVICES.keys())
    print("\nSelect Service:")
    for i, s in enumerate(service_list, 1): print(f"{i}. {s}")
    selected_service = service_list[int(input("Choice #: ")) - 1]

    # 2. Index Requirements
    print("\n--- Indexing Configuration ---")
    idx_out_start = int(input("Start index for Prefix-OUT: "))
    rule_start = int(input("Start index for ACL Rules: "))

    # 3. IP Data
    in_ips = get_in_ips()
    out_ips = SERVICES[selected_service]

    prefix_name_in = f"{selected_telco}_{selected_service}_IN"
    prefix_name_out = f"{selected_telco}_{selected_service}_OUT"
    acl_name = f"{selected_telco}_{selected_service}_EF"

    print("\n" + "=" * 60)
    print(f" FINAL CONFIGURATION: {selected_telco} {selected_service}")
    print("=" * 60)

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
        # Use IPv4Interface to keep the specific IP address provided
        in_iface = ipaddress.IPv4Interface(in_ip_str)
        for out_ip_str in out_ips:
            out_iface = ipaddress.IPv4Interface(out_ip_str)

            # Print the specific IP instead of the network address
            print(f" rule {curr_rule} permit ip source {in_iface.ip} {in_iface.network.hostmask} "
                  f"destination {out_iface.ip} {out_iface.network.hostmask}")
            curr_rule += 5

    print("\n" + "=" * 60)


if __name__ == "__main__":
    main()
