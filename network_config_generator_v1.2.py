import sys
import os
import pandas as pd
import ipaddress
import logging
from datetime import datetime
from colorama import Fore, Style, init

init(autoreset=True)
# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def read_data(file_path):
    """Read Excel data and handle potential errors."""
    try:
        data = pd.read_excel(file_path)
        logger.info(Fore.LIGHTGREEN_EX + f"Data read successfully from {file_path}" + Style.RESET_ALL)
        return data
    except FileNotFoundError:
        logger.error(Fore.LIGHTRED_EX + f"File not found: {file_path}")
    except Exception as e:
        logger.error(Fore.LIGHTRED_EX + f"Error reading Excel file: {e}")
    return None


def validate_data(data):
    """Validate essential columns in the data."""
    required_columns = ['SiteA', 'SiteB', 'LagA', 'LagB', 'Subnet', 'PortType', 'RoutingProto', 'Area']
    for col in required_columns:
        if col not in data.columns:
            logger.error(Fore.LIGHTRED_EX + "Missing required column: %s", col)
            return False
    if (data['RoutingProto'].str.lower() == 'ospf').any() and data['Area'].isnull().any():
        logger.error(Fore.LIGHTRED_EX + "Area column has missing values for OSPF protocols.")
        return False
    return True


def configure_site(data_row, site_name, local_lag, peer_lag, peer_site_name, is_site_a):
    """Generate configuration for a specific site based on data row."""
    config = [
        '#' + 79 * '=',
        f'# On {site_name} ==> {peer_site_name}',
        '#' + 79 * '-',
        'exit all',
        '/config']
    # IP Address formatting
    network = ipaddress.ip_network(data_row.Subnet, strict=False)
    local_ip = network.network_address + (1 if is_site_a else 2)
    peer_ip = network.network_address + (2 if is_site_a else 1)
    # Port configuration
    for n in range(1, count_ports(data_row) + 1):
        port_id = f'port{"A" if is_site_a else "B"}{n}'
        peer_port_id = f'port{"B" if is_site_a else "A"}{n}'
        if port_id in data_row and not pd.isna(data_row[port_id]):
            config.append(generate_port_config(peer_site_name, data_row[port_id], data_row.get(peer_port_id, ''),
                                               data_row.PortType))
    # Lag configuration
    config.extend([
        f'    lag {local_lag}',
        f'        description "To-{peer_site_name}-Lag-{peer_lag}"'
    ])
    for n in range(1, count_ports(data_row) + 1):
        port_id = f'port{"A" if is_site_a else "B"}{n}'
        if port_id in data_row and not pd.isna(data_row[port_id]):
            config.append(f'        port {data_row[port_id]}')
    if data_row.microBFD.lower() == 'yes':
        config.append(generate_mbfd_config(local_ip, peer_ip))
    config.extend([
        '        dynamic-cost',
        '        lacp active',
        '        no shutdown',
        '    exit'
    ])
    # Router configuration
    interface = data_row.get(f'Interface{"A" if is_site_a else "B"}')
    if pd.isna(interface) or not interface:  # Check if InterfaceA or InterfaceB is empty or NaN
        interface = f'To_{peer_site_name}_LAG{peer_lag}'  # Generate an interface
    if len(interface) > 32:
        logger.error(
            Fore.RED + f"Interface '{interface}' is {len(interface)} characters long, which exceeds the 32-character limit.")
    config.append(
        generate_interface_config(interface, f'{local_ip}/{network.prefixlen}', local_lag, peer_lag, peer_site_name,
                                  data_row.BFD))
    # Routing Protocol Configuration
    protocol = data_row.RoutingProto.lower()
    config.append(
        generate_routing_protocol_config(protocol, interface, area=data_row.Area, key=data_row.Auth_Key,
                                         bfd=data_row.BFD))
    # Additional Protocols
    for proto in ['pim', 'mpls', 'rsvp']:
        if data_row[proto].lower() == 'yes':
            config.append(generate_other_protocol_config(proto, interface))
    if data_row.ldp.lower() == 'yes':
        config.append(generate_ldp_config(interface))
    config.append('    exit')
    config.append('exit')
    return config


def count_ports(data_row):
    """Count the existing port columns (e.g., portA1, portB1) in the data row."""
    return len([col for col in data_row.index if col.startswith("portA") or col.startswith("portB")])


def generate_port_config(peer_site_name, port_id, peer_port_id, port_type):
    """Generate configuration for a specific port."""
    port_config_lines = [
        f'    port {port_id}',
        f'        description "To-{peer_site_name}-{port_type}-{peer_port_id}"',
        '        ethernet'
    ]
    if port_type == "GE":
        port_config_lines.append('            autonegotiate limited')
    port_config_lines.extend([
        '            load-balancing-algorithm include-l4',
        '            hold-time up 5',
        '        exit',
        '        no shutdown',
        '    exit'
    ])
    return '\n'.join(port_config_lines)


def generate_mbfd_config(local_ip, peer_ip):
    """Generate BFD configuration."""
    bfd_lines = [
        '        bfd',
        '            family ipv4',
        f'                local-ip-address {local_ip}',
        f'                remote-ip-address {peer_ip}',
        '                no shutdown',
        '            exit',
        '        exit'
    ]
    return '\n'.join(bfd_lines)


def generate_interface_config(interface, address, lag_a, lag_b, site, bfd):
    """Generate router interface configuration."""
    router_lines = [
        '    router',
        f'        interface "{interface}"',
        f'            address {address}',
        f'            description "To-{site}-Lag-{lag_b}"',
        f'            port lag-{lag_a}',
    ]
    if pd.notna(bfd):
        result = bfd.split("/")
        router_lines.append(f'            bfd {result[0]} receive {result[1]} multiplier {result[2]}')
    router_lines.extend([
        '            no shutdown',
        '        exit'
    ])
    return '\n'.join(router_lines)


def generate_routing_protocol_config(protocol, interface, area=None, key=None, bfd=None):
    """Generate OSPF, ISIS, or other protocol configuration."""
    if protocol == "ospf" and area:
        routing_protocol_lines = [
            '        ospf',
            f'            area {area}',
            f'                interface "{interface}"',
            '                     interface-type point-to-point',
        ]
        if pd.notna(key):
            routing_protocol_lines.append(f'                     message-digest-key 10 md5 {key}')
            routing_protocol_lines.append('                     authentication-type message-digest')
        if pd.notna(bfd):
            routing_protocol_lines.append('                     bfd-enable')

        routing_protocol_lines.extend([
            '                     no shutdown',
            '                 exit',
            '            exit',
            '        exit'
        ])
    else:
        routing_protocol_lines = [
            '        isis',
            f'            interface "{interface}"',
            '                level-capability level-2',
            '                interface-type point-to-point',
        ]

        if pd.notna(key):
            routing_protocol_lines.append(f'                hello-authentication-key {key}')
            routing_protocol_lines.append('                hello-authentication-type message-digest')
        if pd.notna(bfd):
            routing_protocol_lines.append('                bfd-enable ipv4')

        routing_protocol_lines.extend([
            '                no shutdown',
            '            exit',
            '        exit'
        ])
    return '\n'.join(routing_protocol_lines)


def generate_ldp_config(interface):
    """Generate LDP configuration."""
    ldp_lines = [
        f'        ldp',
        f'            interface-parameters',
        f'                interface "{interface}"',
        '                    bfd-enable ipv4',
        '                    ipv4',
        '                        no shutdown',
        '                    exit',
        '                    no shutdown',
        '                exit',
        '            exit',
        '        exit'
    ]
    return '\n'.join(ldp_lines)


def generate_other_protocol_config(protocol, interface):
    """Generate PIM configuration."""
    pim_lines = [
        f'        {protocol}',
        f'            interface "{interface}"',
        '                no shutdown',
        '            exit',
        '        exit'
    ]
    return '\n'.join(pim_lines)


def main():
    data_path = "Network_DataPlan.xlsx"
    DataPlan = read_data(data_path)

    if DataPlan is None or not validate_data(DataPlan):
        logger.error(Fore.LIGHTRED_EX + "Exiting due to invalid data.")
        input()
        sys.exit()
    # Determine output preference
    output_choice = input(
        "Save configurations in a single file or multiple files? (Enter s 'single' or m 'multiple'): ").strip().lower()

    output_dir = "FinalConfigFiles"
    os.makedirs(output_dir, exist_ok=True)
    output_config = []
    for _, row in DataPlan.iterrows():
        output_config.extend([
            '#' + 79 * '=',
            f'# Link {row.SiteA} <=> {row.SiteB}',
        ])
        config = configure_site(row, row.SiteA, row.LagA, row.LagB, row.SiteB, is_site_a=True)
        config += configure_site(row, row.SiteB, row.LagB, row.LagA, row.SiteA, is_site_a=False)

        if output_choice == 'm':
            # Save each configuration to an individual file for each site
            output_file = f'FinalConfigFiles/Configuration_{row.SiteA}_to_{row.SiteB}_{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.txt'
            with open(output_file, 'w') as file:
                file.write('\n'.join(config))
            logger.info(Fore.LIGHTGREEN_EX + f"Configuration saved to {output_file}")
        else:
            # Add to single configuration list
            output_config.extend(config)
    if output_choice == 's':
        output_file = f'FinalConfigFiles/Configuration_{datetime.now().strftime("%d-%m-%Y_%H-%M-%S")}.txt'
        with open(output_file, 'w') as file:
            file.write('\n'.join(output_config))
        logger.info(Fore.LIGHTGREEN_EX + f"Configuration saved to {output_file}")

    print(Fore.LIGHTMAGENTA_EX + '\nPress Enter to exit...' + Style.RESET_ALL)
    input()


if __name__ == "__main__":
    main()
