import openpyxl  # liberary to manage the excel
import json


def find_all_device_interfaces(xlf):
    # Read excel file and return a list of dictionaries
    wb = openpyxl.load_workbook(xlf)
    sheet = wb.active
    dev_interfaces = []
    for row in sheet.rows:
        if row[0].row > 1:  # skip first row (column names)
            dev = row[0].value
            role = row[1].value
            interface = row[2].value
            ipaddress = row[3].value
            subnetmask = row[4].value
            dev_interfaces.append({
                "device": dev,
                "role": role,
                "interface": interface,
                "ipaddress": ipaddress,
                "subnetmask": subnetmask
            })
    return dev_interfaces


def make_list_of_devices_and_roles(inventory):
    dev_list = []
    for rec in inventory:
        dev_list.append({
            "dev_name": rec["device"],
            "role": rec["role"]
        })
    return dev_list


def attach_interfaces_to_devices(dev_name, inventory):
    intf_list = []
    for item in inventory:
        if item["device"] == dev_name:
            intf_list.append({
                "interface": item["interface"],
                "ipaddress": item["ipaddress"],
                "subnetmask": item["subnetmask"]
            })

    return intf_list


def main():
    inventory = find_all_device_interfaces("ipdevices.xlsx")
    devices = make_list_of_devices_and_roles(inventory)

    rack_struc = {"rack": []}
    for device in devices:
        dev_dict = {
            "device": {
                # "dev_id": "",
                "dev_name": device["dev_name"],
                "role": device["role"],
                "interfaces": attach_interfaces_to_devices(device["dev_name"], inventory)
            }
        }
        rack_struc["rack"].append(dev_dict)
    return rack_struc


if __name__ == "__main__":
    print(json.dumps(main(), indent=2))
    with open("rack_struc.json", "w") as f:
        json.dump(main(), f, indent=2)
