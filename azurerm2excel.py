import datetime
import json
import os
import re
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def parse_attributes(attributes, parent_key=""):
    attribute_list = []
    for key, value in attributes.items():
        if isinstance(value, dict):
            attribute_list.extend(
                parse_attributes(value, f"{parent_key}{key}.")
            )  # Recursive call for nested dictionaries
        elif isinstance(value, list):
            for i, item in enumerate(value):
                if isinstance(item, dict):
                    attribute_list.extend(
                        parse_attributes(item, f"{parent_key}{key}[{i}].")
                    )  # Recursive call for nested dictionaries in lists
                else:
                    attribute_list.append(
                        (f"{parent_key}{key}[{i}]", item)
                    )  # Add list item as a separate attribute
        else:
            attribute_list.append(
                (f"{parent_key}{key}", value)
            )  # Add simple key-value pair
    return attribute_list


def load_descriptions(
    description_folders,
):  # description_folder を description_folders に変更
    descriptions = {}
    for description_folder in description_folders:  # 各フォルダをループ
        for filename in os.listdir(description_folder):
            if filename.endswith(".json"):
                type_name = filename.split(".")[0]
                with open(
                    os.path.join(description_folder, filename), "r", encoding="utf-8"
                ) as f:
                    descriptions[type_name] = json.load(f)
    return descriptions


def apply_styles(
    ws, header_fill, header_font, default_font, left_alignment, thin_border
):
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    for row in ws.iter_rows(min_row=2):  # Skip header row
        for cell in row:
            cell.font = default_font
            cell.alignment = left_alignment
            cell.border = thin_border
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width


def write_to_excel(resources_by_type, descriptions, output_folder):
    for resource_type, resources in resources_by_type.items():
        wb = Workbook()
        for resource in resources:
            # Limit sheet title to 30 characters
            sheet_title = resource["name"][:30]
            ws = wb.create_sheet(title=sheet_title)
            attribute_list = []
            for instance in resource["instances"]:
                attribute_list.extend(parse_attributes(instance["attributes"]))

            # Add header row
            header = ["Arguments", "Value", "Description"]
            ws.append(header)

            # Apply background color, font color, and font to the header row
            header_fill = PatternFill(
                start_color="0000FF", end_color="0000FF", fill_type="solid"
            )  # Blue background
            header_font = Font(
                color="FFFFFF", name="Yu Gothic UI", bold=True
            )  # White text, Yu Gothic UI font, bold
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font

            if resource_type == "azurerm_network_security_group":
                security_rules = [
                    attribute
                    for attribute in attribute_list
                    if "security_rule" in attribute[0]
                ]
                other_attributes = [
                    attribute
                    for attribute in attribute_list
                    # if "security_rule" not in attribute[0]
                ]

                # security_ruleだけ別のシートに出力
                if security_rules:
                    ws_security_rule = wb.create_sheet(title=f"{sheet_title}_rule")
                    # Add header row for security_rule sheet
                    ws_security_rule.append(
                        [
                            "Rule Index",
                            "Direction",
                            "Priority",
                            "Access",
                            "Name",
                            "Description",
                            "Destination Address Prefix",
                            "Destination Port Range",
                            "Protocol",
                            "Source Address Prefix",
                            "Source Port Range",
                        ]
                    )

                    rule_dict = defaultdict(dict)
                    for rule in security_rules:
                        rule_key = re.sub(r"\[\d+\]", "", rule[0])
                        rule_index = re.search(r"\[(\d+)\]", rule[0]).group(1)
                        if rule_key in rule_dict[rule_index]:
                            rule_dict[rule_index][rule_key] = str(rule_dict[rule_index][rule_key]) + f"\n{rule[1]}"
                        else:
                            rule_dict[rule_index][rule_key] = rule[1]

                    # Sort rules by priority
                    sorted_rules = sorted(
                        rule_dict.items(),
                        key=lambda x: (
                            x[1].get("security_rule.direction", ""),
                            int(x[1].get("security_rule.priority", 0)),
                        ),
                    )

                    for rule_index, rule_attrs in sorted_rules:
                        row = [f"security_rule[{rule_index}]"]
                        for key in [
                            "security_rule.direction",
                            "security_rule.priority",
                            "security_rule.access",
                            "security_rule.name",
                            "security_rule.description",
                            "security_rule.destination_address_prefix",
                            "security_rule.destination_port_range",
                            "security_rule.protocol",
                            "security_rule.source_address_prefix",
                            "security_rule.source_port_range",
                        ]:
                            row.append(str(rule_attrs.get(key, "")))
                        ws_security_rule.append(row)

                    # Apply styles to the new sheet
                    apply_styles(
                        ws_security_rule,
                        header_fill,
                        header_font,
                        default_font,
                        left_alignment,
                        thin_border,
                    )

                # Add other attributes to the original sheet
                for attribute in other_attributes:
                    attribute_path = attribute[0]
                    normalized_path = re.sub(r"\[\d+\]", "", attribute_path)
                    description = descriptions.get(resource_type, {}).get(
                        normalized_path, ""
                    )
                    value = str(attribute[1])
                    ws.append([attribute[0], value, description])
            elif resource_type == "azurerm_firewall_policy_rule_collection_group":
                application_rule_collections = [
                    attribute
                    for attribute in attribute_list
                    if "application_rule_collection" in attribute[0]
                ]
                network_rule_collections = [
                    attribute
                    for attribute in attribute_list
                    if "network_rule_collection" in attribute[0]
                ]
                nat_rule_collections = [
                    attribute
                    for attribute in attribute_list
                    if "nat_rule_collection" in attribute[0]
                ]
                other_rule_collections = [
                    attribute
                    for attribute in attribute_list
                    # if "application_rule_collection" not in attribute[0]
                    # and "network_rule_collection" not in attribute[0]
                    # and "nat_rule_collection" not in attribute[0]
                ]

                # application_rule_collectionだけ別のシートに出力
                if application_rule_collections:
                    ws_appcol = wb.create_sheet(title=f"{sheet_title}_appcol")
                    ws_appcolrule = wb.create_sheet(title=f"{sheet_title}_appcolrule")
                    # Add header row for application_rule_collection sheet
                    ws_appcol.append(
                        [
                            "Collection Index",
                            "Name",
                            "Priority",
                            "Action",
                        ]
                    )
                    ws_appcolrule.append(
                        [
                            "Rule Index(=Collection Index_Rule Index)",
                            "Name",
                            "Description",
                            "Destination Addresses",
                            "Destination FQDN Tags",
                            "Destination FQDNs",
                            "Destination URLs",
                            "Protocols Port",
                            "Protocols Type",
                            "Source Addresses",
                            "Source IP Groups",
                            "Terminate TLS",
                            "Web Categories",
                        ]
                    )

                    appcol_dict = defaultdict(dict)
                    appcolrule_dict = defaultdict(dict)
                    for app_col in application_rule_collections:
                        app_col_key = re.sub(r"\[\d+\]", "", app_col[0])
                        app_col_index = re.search(r"\[(\d+)\]", app_col[0]).group(1)
                        appcol_dict[app_col_index][app_col_key] = app_col[1]
                        if "rule" in app_col[0]:
                            match = re.search(r"rule\[(\d+)\]", app_col[0])
                            if match:
                                app_col_rule_index = match.group(1)
                                if f"{app_col_index}_{app_col_rule_index}" in appcolrule_dict:
                                    if app_col_key in appcolrule_dict[f"{app_col_index}_{app_col_rule_index}"]:
                                        appcolrule_dict[f"{app_col_index}_{app_col_rule_index}"][app_col_key] = str(appcolrule_dict[f"{app_col_index}_{app_col_rule_index}"].get(app_col_key, "")) + f"\n{app_col[1]}"
                                    else:
                                        appcolrule_dict[f"{app_col_index}_{app_col_rule_index}"][app_col_key] = app_col[1]
                                else:
                                    appcolrule_dict[f"{app_col_index}_{app_col_rule_index}"][app_col_key] = app_col[1]
                            else:
                                app_col_rule_index = None
                    # Sort app_cols by priority
                    sorted_appcols = sorted(
                        appcol_dict.items(),
                        key=lambda x: int(
                            x[1].get("application_rule_collection.priority", 0)
                        ),
                    )

                    for app_col_index, app_col_attrs in sorted_appcols:
                        row = [f"application_rule_collection[{app_col_index}]"]
                        for key in [
                            "application_rule_collection.name",
                            "application_rule_collection.priority",
                            "application_rule_collection.action",
                        ]:
                            row.append(str(app_col_attrs.get(key, "")))
                        ws_appcol.append(row)

                    for (
                        app_col_rule_index,
                        app_col_rule_attrs,
                    ) in appcolrule_dict.items():
                        row = [
                            f"application_rule_collection.rule[{app_col_rule_index}]"
                        ]
                        for key in [
                            "application_rule_collection.rule.name",
                            "application_rule_collection.rule.description",
                            "application_rule_collection.rule.destination_addresses",
                            "application_rule_collection.rule.destination_fqdn_tags",
                            "application_rule_collection.rule.destination_fqdns",
                            "application_rule_collection.rule.destination_urls",
                            "application_rule_collection.rule.protocols.port",
                            "application_rule_collection.rule.protocols.type",
                            "application_rule_collection.rule.source_addresses",
                            "application_rule_collection.rule.source_ip_groups",
                            "application_rule_collection.rule.terminate_tls",
                            "application_rule_collection.rule.web_categories",
                        ]:
                            row.append(str(app_col_rule_attrs.get(key, "")))
                        ws_appcolrule.append(row)

                    # Apply styles to the new sheet
                    apply_styles(
                        ws_appcol,
                        header_fill,
                        header_font,
                        default_font,
                        left_alignment,
                        thin_border,
                    )
                    apply_styles(
                        ws_appcolrule,
                        header_fill,
                        header_font,
                        default_font,
                        left_alignment,
                        thin_border,
                    )

                # network_rule_collectionだけ別のシートに出力
                if network_rule_collections:
                    ws_netcol = wb.create_sheet(title=f"{sheet_title}_netcol")
                    ws_netcolrule = wb.create_sheet(title=f"{sheet_title}_netcolrule")
                    # Add header row for network_rule_collection sheet
                    ws_netcol.append(
                        [
                            "Collection Index",
                            "Name",
                            "Priority",
                            "Action",
                        ]
                    )
                    ws_netcolrule.append(
                        [
                            "Rule Index(=Collection Index_Rule Index)",
                            "Name",
                            "Description",
                            "Destination Addresses",
                            "Destination FQDNs",
                            "Destination IP Groups",
                            "Destination Ports",
                            "Protocols",
                            "Source Addresses",
                            "Source IP Groups",
                        ]
                    )

                    netcol_dict = defaultdict(dict)
                    netcolrule_dict = defaultdict(dict)
                    for net_col in network_rule_collections:
                        net_col_key = re.sub(r"\[\d+\]", "", net_col[0])
                        net_col_index = re.search(r"\[(\d+)\]", net_col[0]).group(1)
                        netcol_dict[net_col_index][net_col_key] = net_col[1]
                        if "rule" in net_col[0]:
                            match = re.search(r"rule\[(\d+)\]", net_col[0])
                            if match:
                                net_col_rule_index = match.group(1)
                                if f"{net_col_index}_{net_col_rule_index}" in netcolrule_dict:
                                    if net_col_key in netcolrule_dict[f"{net_col_index}_{net_col_rule_index}"]:
                                        netcolrule_dict[f"{net_col_index}_{net_col_rule_index}"][net_col_key] = str(netcolrule_dict[f"{net_col_index}_{net_col_rule_index}"][net_col_key]) + f"\n{net_col[1]}"
                                    else:
                                        netcolrule_dict[f"{net_col_index}_{net_col_rule_index}"][net_col_key] = net_col[1]
                                else:
                                    netcolrule_dict[f"{net_col_index}_{net_col_rule_index}"][net_col_key] = net_col[1]
                            else:
                                net_col_rule_index = None

                    # Sort net_cols by priority
                    sorted_netcols = sorted(
                        netcol_dict.items(),
                        key=lambda x: int(
                            x[1].get("network_rule_collection.priority", 0)
                        ),
                    )

                    for net_col_index, net_col_attrs in sorted_netcols:
                        row = [f"network_rule_collection[{net_col_index}]"]
                        for key in [
                            "network_rule_collection.name",
                            "network_rule_collection.priority",
                            "network_rule_collection.action",
                        ]:
                            row.append(str(net_col_attrs.get(key, "")))
                        ws_netcol.append(row)

                    for (
                        net_col_rule_index,
                        net_col_rule_attrs,
                    ) in netcolrule_dict.items():
                        row = [
                            f"net_rule_collection.rule[{net_col_rule_index}]"
                        ]
                        for key in [
                            "network_rule_collection.rule.name",
                            "network_rule_collection.rule.description",
                            "network_rule_collection.rule.destination_addresses",
                            "network_rule_collection.rule.destination_fqdns",
                            "network_rule_collection.rule.destination_ip_groups",
                            "network_rule_collection.rule.destination_ports",
                            "network_rule_collection.rule.protocols",
                            "network_rule_collection.rule.source_addresses",
                            "network_rule_collection.rule.source_ip_groups",
                        ]:
                            row.append(str(net_col_rule_attrs.get(key, "")))
                        ws_netcolrule.append(row)

                    # Apply styles to the new sheet
                    apply_styles(
                        ws_netcol,
                        header_fill,
                        header_font,
                        default_font,
                        left_alignment,
                        thin_border,
                    )
                    apply_styles(
                        ws_netcolrule,
                        header_fill,
                        header_font,
                        default_font,
                        left_alignment,
                        thin_border,
                    )
                
                # nat_rule_collectionだけ別のシートに出力
                if nat_rule_collections:
                    ws_natcol = wb.create_sheet(title=f"{sheet_title}_natcol")
                    ws_natcolrule = wb.create_sheet(title=f"{sheet_title}_natcolrule")
                    # Add header row for nat_rule_collection sheet
                    ws_natcol.append(
                        [
                            "Collection Index",
                            "Name",
                            "Priority",
                            "Action",
                        ]
                    )
                    ws_natcolrule.append(
                        [
                            "Rule Index(=Collection Index_Rule Index)",
                            "Name",
                            "Description",
                            "Destination Address",
                            "Destination Ports",
                            "Protocols",
                            "Source Addresses",
                            "Source IP Groups",
                            "Translated Address",
                            "Translated FQDN",
                            "Translated Port",
                        ]
                    )

                    natcol_dict = defaultdict(dict)
                    natcolrule_dict = defaultdict(dict)
                    for nat_col in nat_rule_collections:
                        nat_col_key = re.sub(r"\[\d+\]", "", nat_col[0])
                        nat_col_index = re.search(r"\[(\d+)\]", nat_col[0]).group(1)
                        natcol_dict[nat_col_index][nat_col_key] = nat_col[1]
                        if "rule" in nat_col[0]:
                            match = re.search(r"rule\[(\d+)\]", nat_col[0])
                            if match:
                                nat_col_rule_index = match.group(1)
                                if f"{nat_col_index}_{nat_col_rule_index}" in natcolrule_dict:
                                    if nat_col_key in natcolrule_dict[f"{nat_col_index}_{nat_col_rule_index}"]:
                                        natcolrule_dict[f"{nat_col_index}_{nat_col_rule_index}"][nat_col_key] = str(natcolrule_dict[f"{nat_col_index}_{nat_col_rule_index}"][nat_col_key]) + f"\n{nat_col[1]}"
                                    else:
                                        natcolrule_dict[f"{nat_col_index}_{nat_col_rule_index}"][nat_col_key] = nat_col[1]
                                else:
                                    natcolrule_dict[f"{nat_col_index}_{nat_col_rule_index}"][nat_col_key] = nat_col[1]
                            else:
                                nat_col_rule_index = None

                    # Sort nat_cols by priority
                    sorted_natcols = sorted(
                        natcol_dict.items(),
                        key=lambda x: int(
                            x[1].get("nat_rule_collection.priority", 0)
                        ),
                    )

                    for nat_col_index, nat_col_attrs in sorted_natcols:
                        row = [f"nat_rule_collection[{nat_col_index}]"]
                        for key in [
                            "nat_rule_collection.name",
                            "nat_rule_collection.priority",
                            "nat_rule_collection.action",
                        ]:
                            row.append(str(nat_col_attrs.get(key, "")))
                        ws_natcol.append(row)

                    for (
                        nat_col_rule_index,
                        nat_col_rule_attrs,
                    ) in natcolrule_dict.items():
                        row = [
                            f"nat_rule_collection.rule[{nat_col_rule_index}]"
                        ]
                        for key in [
                            "nat_rule_collection.rule.name",
                            "nat_rule_collection.rule.description",
                            "nat_rule_collection.rule.destination_address",
                            "nat_rule_collection.rule.destination_ports",
                            "nat_rule_collection.rule.protocols",
                            "nat_rule_collection.rule.source_addresses",
                            "nat_rule_collection.rule.source_ip_groups",
                            "nat_rule_collection.rule.translated_address",
                            "nat_rule_collection.rule.translated_fqdn",
                            "nat_rule_collection.rule.translated_port",
                        ]:
                            row.append(str(nat_col_rule_attrs.get(key, "")))
                        ws_natcolrule.append(row)

                    # Apply styles to the new sheet
                    apply_styles(
                        ws_natcol,
                        header_fill,
                        header_font,
                        default_font,
                        left_alignment,
                        thin_border,
                    )
                    apply_styles(                 
                        ws_natcolrule,
                        header_fill,
                        header_font,
                        default_font,
                        left_alignment,
                        thin_border,
                    )

                for attribute in other_rule_collections:
                    attribute_path = attribute[0]
                    normalized_path = re.sub(r"\[\d+\]", "", attribute_path)
                    description = descriptions.get(resource_type, {}).get(
                        normalized_path, ""
                    )
                    value = str(attribute[1])
                    ws.append([attribute[0], value, description])

            else:
                for attribute in attribute_list:
                    attribute_path = attribute[0]
                    normalized_path = re.sub(r"\[\d+\]", "", attribute_path)
                    description = descriptions.get(resource_type, {}).get(
                        normalized_path, ""
                    )
                    value = str(attribute[1])
                    ws.append([attribute[0], value, description])

            # Apply font and alignment to all other cells
            default_font = Font(name="Yu Gothic UI")
            left_alignment = Alignment(
                horizontal="left", vertical="top", wrap_text=True
            )  # Left alignment with text wrapping and top alignment
            for row in ws.iter_rows(min_row=2):  # Skip header row
                for cell in row:
                    cell.font = default_font
                    cell.alignment = left_alignment

            # Apply border to all cells
            thin_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = thin_border

            # Adjust column widths
            ws.column_dimensions["B"].width = 100  # Set width for Value column
            ws.column_dimensions["C"].width = 100  # Set width for Description column

            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter  # Get the column name
                if column not in ["B", "C"]:  # Skip Value and Description columns
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = max_length + 2
                    ws.column_dimensions[column].width = adjusted_width

        # Remove the default sheet created by openpyxl
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        output_path = os.path.join(output_folder, f"{resource_type}.xlsx")
        wb.save(output_path)
        print(f"Excel file saved to {output_path}")


def process_tfstate(tfstate_file, description_folders, output_folder):
    with open(tfstate_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    resources_by_type = defaultdict(list)
    for res in data["resources"]:
        if res["mode"] == "managed":
            resources_by_type[res["type"]].append(res)

    descriptions = load_descriptions(description_folders)
    write_to_excel(resources_by_type, descriptions, output_folder)


if __name__ == "__main__":
    if len(sys.argv) < 3:  # 引数の数を変更
        print(
            "Usage: python azurerm2excel.py <path_to_tfstate_file> <description_folder1> [<description_folder2> ...]"
        )  # メッセージを変更
        sys.exit(1)

    tfstate_file = sys.argv[1]
    description_folders = sys.argv[
        2:
    ]  # description_folder を description_folders に変更
    output_folder = "terraoutput"  # 固定の出力フォルダ

    if not os.path.isfile(tfstate_file):
        print(f"Error: The file {tfstate_file} does not exist。")
        sys.exit(1)

    for description_folder in description_folders:  # 各フォルダをチェック
        if not os.path.isdir(description_folder):
            print(f"Error: The directory {description_folder} does not exist。")
            sys.exit(1)

    # フォルダが存在しない場合は作成する
    if not os.path.isdir(output_folder):
        print(f"Creating output folder: {output_folder}")
        os.makedirs(output_folder)

    # YYYYMMDDss フォルダを作成
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    output_folder = os.path.join(output_folder, timestamp)
    os.makedirs(output_folder)

    process_tfstate(
        tfstate_file, description_folders, output_folder
    )  # description_folder を description_folders に変更
