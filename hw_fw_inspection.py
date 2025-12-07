from pathlib import Path
import tarfile as tf
import openpyxl as opx
import pandas as pd
import re
import ipaddress

# 사용자가 변경해야할 값들
folder = Path(r"C:\Users\lionm\Downloads\LogCollectResult\20251107162128\device_logs")
output_file = Path(r"C:\Users\lionm\Downloads\output_one_click_log.xlsx")

bios_config_dict_k1_k2_k3_k4_sa = {
    "PcieAspmControl": "Disabled",
    "SriovEnablePolicy": "Disabled",
    "ProcessorHyperThreadingDisable": "ALL LPs",
    "ProcessorVmxEnable": "Disabled",
    "VTdSupport": "Disabled",
    "ProcessorX2apic": "Disabled",
    "NumaEn": "Enabled",
    "SncEn": "Disabled",
    "VMDConfigEnable": "Disabled",
    "BenchMarkSelection": "Custom",
    "StaticTurbo": "Disabled",
    "PowerSaving": "Disabled",
    "ProcessorEistEnable": "Disabled",
    "EETurboDisable": "Disabled",
    "ProcessorFlexibleRatioOverrideEnable": "Disabled",
    "ProcessorHWPMEnable": "Disabled",
    "MonitorMWait": "Disabled",
    "C6Enable": "Disabled",
    "ProcessorC1eEnable": "Disabled",
    "PackageCState": "C0/C1 State",
    "PwrPerfTuning": "BIOS Controls EPB",
    "AltEngPerfBIAS": "Performance",
    "OptimizedPowerMode": "Disabled",
    "BootType": "UEFIBoot",
    "HttpNetworkProtocol": "Disabled",
    "UsbBoot": "Disabled",
    "SyncOrder": "Disabled",
    "PciePortEnable1": "Auto",
    "PciePortEnable9": "Auto",
    "BootTypeOrder0": "HardDiskDrive",
    "BootTypeOrder1": "PXE",
    "BootTypeOrder2": "DVDROMDrive",
    "BootTypeOrder3": "Others"
}

def main():

    df_data = []

    # 원클릭로그 폴더 내의 모든 .tar.gz 파일을 순회
    for archived_filename in folder.glob("*.tar.gz"):
        if archived_filename.is_file():
            with tf.open(archived_filename, "r:gz") as tar:

                # ipinfo_info >> bmc ip 정보 추출
                ipinfo_info_tarinfo = tar.getmember("dump_info/RTOSDump/networkinfo/ipinfo_info")
                if ipinfo_info_tarinfo.isfile():
                    with tar.extractfile(ipinfo_info_tarinfo) as ipinfo_info_file:
                        ipinfo_info_content = ipinfo_info_file.read().decode('utf-8')
                        bmc_ip_search_by_regex = re.search(r'IP Address\s+:\s+([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)', ipinfo_info_content, re.MULTILINE)
                        bmc_ip = bmc_ip_search_by_regex.group(1) if bmc_ip_search_by_regex else "regex extraction failed"

                # ipinfo_info 파일 자체를 찾지 못한 경우
                else:
                    bmc_ip = "no ipinfo_info file"

                # uname_info >> sn 정보 추출
                uname_info_tarinfo = tar.getmember("dump_info/RTOSDump/sysinfo/uname_info")
                if uname_info_tarinfo.isfile():
                    with tar.extractfile(uname_info_tarinfo) as uname_info_file:
                        uname_info_content = uname_info_file.read().decode('utf-8')
                        sn = uname_info_content.split()[1]

                # uname_info 파일 자체를 찾지 못한 경우
                else:
                    sn = "no uname_info file"

                # server_config.txt >> cpu, memory, nvme, ssd, hdd, raid, psu, gpu 정보 추출
                server_config_txt_tarinfo = tar.getmember("dump_info/RTOSDump/versioninfo/server_config.txt")
                if server_config_txt_tarinfo.isfile():
                    with tar.extractfile(server_config_txt_tarinfo) as server_config_txt_file:
                        server_config_txt_content = server_config_txt_file.read().decode('utf-8')

                        # CPU 정보 추출
                        cpu_count = 0
                        cpu_info_regex = re.search(r"(-+Cpu\s*info.*)(?=\n\s*\n-*Memory\s*info)", server_config_txt_content, re.DOTALL)
                        if cpu_info_regex:
                            cpu_info = cpu_info_regex.group(1)
                            valid_cpu_info = [line for line in cpu_info.splitlines()[2:]]

                            cpu_sn = [line.split(',')[-1].lstrip() for line in valid_cpu_info]
                            cpu_vendor = [line.split()[4] for line in valid_cpu_info]
                            cpu_model = [" ".join(line.split()[5:8]).rstrip(',') for line in valid_cpu_info]
                            cpu_core = [line.split(',')[4].lstrip() for line in valid_cpu_info]
                            cpu_thread = [line.split(',')[5].lstrip() for line in valid_cpu_info]
                            cpu_count = len(valid_cpu_info)

                        # cpu info 섹션 자체를 찾지 못한 경우
                        else:
                            cpu_sn = "cpu info not found in server_config.txt"
                            cpu_vendor = "cpu info not found in server_config.txt"
                            cpu_model = "cpu info not found in server_config.txt"
                            cpu_core = "cpu info not found in server_config.txt"
                            cpu_thread = "cpu info not found in server_config.txt"
                            cpu_count = "cpu info not found in server_config.txt"

                        # Memory 정보 추출
                        memory_info_regex = re.search(r"(-+Memory\s*info-+.*)(?=\n\s*\n-*Card info)", server_config_txt_content, re.DOTALL)
                        if memory_info_regex:
                            memory_info = memory_info_regex.group(1)
                            valid_memory_info = [line for line in memory_info.splitlines()[2:] if line.split()[8] != "Unknown,"]

                            memory_sn = [line.split(',')[8].lstrip() for line in valid_memory_info]
                            memory_vendor = [line.split(',')[3].lstrip() for line in valid_memory_info]
                            memory_model = [line.split(',')[14].lstrip().rstrip(',') for line in valid_memory_info]
                            memory_type = [line.split(',')[7].lstrip() for line in valid_memory_info]
                            memory_maximum_speed = [line.split(',')[5].lstrip() for line in valid_memory_info]
                            memory_configured_speed = [line.split(',')[6].lstrip() for line in valid_memory_info]
                            memory_capacity = [line.split(',')[4].lstrip() for line in valid_memory_info]
                            memory_count = len(valid_memory_info)
                        else:
                            memory_sn = "memory info not found in server_config.txt"
                            memory_vendor = "memory info not found in server_config.txt"
                            memory_model = "memory info not found in server_config.txt"
                            memory_type = "memory info not found in server_config.txt"
                            memory_maximum_speed = "memory info not found in server_config.txt"
                            memory_configured_speed = "memory info not found in server_config.txt"
                            memory_capacity = "memory info not found in server_config.txt"
                            memory_count = "memory info not found in server_config.txt"

                        # NVMe 정보 추출
                        nvme_info_regex = re.search(r"(Pass Through Drives Information.*)", server_config_txt_content, re.DOTALL)
                        if nvme_info_regex:
                            nvme_info = nvme_info_regex.group(1)
                            nvme_sn_list = re.findall(r"Serial Number\s*:\s*[^\n\r]+", nvme_info)
                            nvme_sn = [sn.split(":")[1].lstrip().rstrip() for sn in nvme_sn_list]
                            nvme_vendor_list = re.findall(r"Manufacturer\s*:\s*[^\n\r]+", nvme_info)
                            nvme_vendor = [sn.split(":")[1].lstrip().rstrip() for sn in nvme_vendor_list]
                            nvme_model_list = re.findall(r"Model\s*:\s*[^\n\r]+", nvme_info)
                            nvme_model = [sn.split(":")[1].lstrip().rstrip() for sn in nvme_model_list]
                            nvme_interface_list = re.findall(r"Interface\s*Type\s*:\s*[^\n\r]+", nvme_info)
                            nvme_interface = [sn.split(":")[1].lstrip().rstrip() for sn in nvme_interface_list]
                            nvme_capacity_list = re.findall(r"Capacity\s*:\s*[^\n\r]+", nvme_info)
                            nvme_capacity = [sn.split(":")[1].lstrip().rstrip() for sn in nvme_capacity_list]
                            nvme_fw_list = re.findall(r"Firmware\s*Version\s*:\s*[^\n\r]+", nvme_info)
                            nvme_fw = [sn.split(":")[1].lstrip().rstrip() for sn in nvme_fw_list]
                            nvme_id_list = re.findall(r"ID\s*:\s*[^\n\r]+", nvme_info)
                            nvme_count = len(nvme_id_list)

                        # nvme info 섹션 자체를 찾지 못한 경우
                        else:
                            nvme_sn = "nvme info not found in server_config.txt"
                            nvme_vendor = "nvme info not found in server_config.txt"
                            nvme_model = "nvme info not found in server_config.txt"
                            nvme_interface = "nvme info not found in server_config.txt"
                            nvme_capacity = "nvme info not found in server_config.txt"
                            nvme_fw = "nvme info not found in server_config.txt"
                            nvme_count = "nvme info not found in server_config.txt"

                        # PSU 정보 추출
                        psu_info_regex = re.search(r"(-+PSU\s*info.*)(?=\n\s*\n-*Storage\s*info)", server_config_txt_content, re.DOTALL)
                        if psu_info_regex:
                            psu_info = psu_info_regex.group(1)
                            valid_psu_info = [line for line in psu_info.splitlines()[2:]]
                            psu_sn = [line.split('|')[4].lstrip().rstrip() for line in valid_psu_info]
                            psu_capacity = [line.split('|')[6].lstrip().rstrip() for line in valid_psu_info]
                            psu_count = len(valid_psu_info)

                        # psu info 섹션 자체를 찾지 못한 경우
                        else:
                            psu_sn = "psu info not found in server_config.txt"
                            psu_capacity = "psu info not found in server_config.txt"
                            psu_count = "psu info not found in server_config.txt"

                # server_config.txt 파일 자체를 찾지 못한 경우
                else:
                    cpu_sn = "no server_config.txt file"
                    cpu_vendor = "no server_config.txt file"
                    cpu_model = "no server_config.txt file"
                    cpu_core = "no server_config.txt file"
                    cpu_thread = "no server_config.txt file"
                    cpu_count = "no server_config.txt file"

                    memory_sn = "no server_config.txt file"
                    memory_vendor = "no server_config.txt file"
                    memory_model = "no server_config.txt file"
                    memory_type = "no server_config.txt file"
                    memory_maximum_speed = "no server_config.txt file"
                    memory_configured_speed = "no server_config.txt file"
                    memory_capacity = "no server_config.txt file"
                    memory_count = "no server_config.txt file"

                    psu_sn = "no server_config.txt file"
                    psu_capacity = "no server_config.txt file"
                    psu_count = "no server_config.txt file"

                # netcard_info.txt >> nic 정보 추출
                netcard_info_tarinfo = tar.getmember("dump_info/LogDump/netcard/netcard_info.txt")
                if netcard_info_tarinfo.isfile():
                    with tar.extractfile(netcard_info_tarinfo) as netcard_info_file:
                        netcard_info_content = netcard_info_file.read().decode('utf-8')
                        ocp1_model_match_list = re.findall(r'ProductName\s*:([^\n\r]+)', netcard_info_content, re.MULTILINE)
                        ocp1_model = ocp1_model_match_list[-1] if ocp1_model_match_list else "not found in netcard_info.txt"
                        ocp1_vendor_match_list = re.findall(r'Manufacture\s*:([^\n\r]+)', netcard_info_content, re.MULTILINE)
                        ocp1_vendor = ocp1_vendor_match_list[-1] if ocp1_vendor_match_list else "not found in netcard_info.txt"
                        ocp1_mac_match_list = re.findall(r'^\s*MacAddr:([^\n\r]+)', netcard_info_content, re.MULTILINE)
                        ocp1_mac = ocp1_mac_match_list[-2] if ocp1_mac_match_list else "not found in netcard_info.txt"
                        ocp1_slot_match_list = re.findall(r'SlotId\s*:([^\n\r]+)', netcard_info_content, re.MULTILINE)
                        ocp1_slot = ocp1_slot_match_list[-1] if ocp1_slot_match_list else "not found in netcard_info.txt"

                # netcard_info.txt 파일 자체를 찾지 못한 경우
                else:
                    ocp1_model = "no netcard_info.txt file"
                    ocp1_vendor = "no netcard_info.txt file"
                    ocp1_mac = "no netcard_info.txt file"
                    ocp1_slot = "no netcard_info.txt file"

                # app_revision.txt >> bios, bmc, cpld, vrd, pn 정보 추출
                app_revision_txt_tarinfo = tar.getmember("dump_info/RTOSDump/versioninfo/app_revision.txt")
                if app_revision_txt_tarinfo.isfile():
                    with tar.extractfile(app_revision_txt_tarinfo) as app_revision_txt_file:
                        app_revision_txt_content = app_revision_txt_file.read().decode('utf-8')
                        pn_regex = re.search(r'Product\s*Name:\s*([^\n\r]+)', app_revision_txt_content)
                        pn = pn_regex.group(1) if pn_regex else "not found in app_revision.txt"
                        bios_version_regex = re.search(r'Active BIOS\s+Version\s*:\s*([^\n\r]+)', app_revision_txt_content)
                        bios_version = bios_version_regex.group(1) if bios_version_regex else "not found in app_revision.txt"
                        bmc_version_regex = re.search(r'Active iBMC\s+Version\s*:\s*([^\n\r]+)', app_revision_txt_content)
                        bmc_version = bmc_version_regex.group(1) if bmc_version_regex else "not found in app_revision.txt"
                        mb_cpld_version_regex = re.search(r'CPLD\s+Version\s*:\s*([^\n\r]+)', app_revision_txt_content)
                        mb_cpld_version = mb_cpld_version_regex.group(1) if mb_cpld_version_regex else "not found in app_revision.txt"
                        bp1_cpld_version_regex = re.search(r'Disk BP1\s+CPLD\s+Version\s*:\s*([^\n\r]+)', app_revision_txt_content)
                        bp1_cpld_version = bp1_cpld_version_regex.group(1) if bp1_cpld_version_regex else "not found in app_revision.txt"
                        bp2_cpld_version_regex = re.search(r'Disk BP2\s+CPLD\s+Version\s*:\s*([^\n\r]+)', app_revision_txt_content)
                        bp2_cpld_version = bp2_cpld_version_regex.group(1) if bp2_cpld_version_regex else "not found in app_revision.txt"

                # app_revision.txt 파일 자체를 찾지 못한 경우
                else:
                    pn = "no app_revision.txt file"
                    bios_version = "no app_revision.txt file"
                    bmc_version = "no app_revision.txt file"
                    mb_cpld_version = "no app_revision.txt file"
                    bp1_cpld_version = "no app_revision.txt file"
                    bp2_cpld_version = "no app_revision.txt file"

                # BIOS/currentvalue.json >> bios 설정 정보 추출
                bios_currentvalue_json_tarinfo = tar.getmember("dump_info/AppDump/BIOS/currentvalue.json")
                if bios_currentvalue_json_tarinfo.isfile():
                    with tar.extractfile(bios_currentvalue_json_tarinfo) as bios_currentvalue_json_file:
                        bios_currentvalue_json_content = bios_currentvalue_json_file.read().decode('utf-8')
                        bios_mismatch_list = []
                        bios_mismatch_count = 0
                        for key, expected_value in bios_config_dict_k1_k2_k3_k4_sa.items():
                            bios_actual_value_regex = re.search(rf'"{key}"\s*:\s*"([^"]+)"', bios_currentvalue_json_content)
                            if bios_actual_value_regex:
                                bios_actual_value = bios_actual_value_regex.group(1)
                                if bios_actual_value != expected_value:
                                    bios_mismatch_list.append(key)
                                    bios_mismatch_count += 1
                            else:
                                bios_mismatch_list.append(key)
                                bios_mismatch_count += 1
                        bios_configuration_status = True if bios_mismatch_count == 0 else False
                        bios_ocp1_regex = re.search(rf'"PciePortEnable1"\s*:\s*"([^"]+)"', bios_currentvalue_json_content)
                        bios_ocp1_status = bios_ocp1_regex.group(1) if bios_ocp1_regex else "not found in currentvalue.json"
                        bios_disable_boot_drive_status_regex = re.search(rf'"PciePortEnable9"\s*:\s*"([^"]+)"', bios_currentvalue_json_content)
                        bios_disable_boot_drive_status = bios_disable_boot_drive_status_regex.group(1) if bios_disable_boot_drive_status_regex else "not found in currentvalue.json"
                else:
                    bios_mismatch_list = "no currentvalue.json file"
                    bios_mismatch_count = "no currentvalue.json file"
                    bios_configuration_status = "no currentvalue.json file"
                    bios_ocp1_status = "no currentvalue.json file"
                    bios_disable_boot_drive_status = "no currentvalue.json file"

        # 압축파일이 파일로 인식되지 않은 경우
        else:
            bmc_ip = "not a file"
            sn = "not a file"

            cpu_sn = "not a file"
            cpu_vendor = "not a file"
            cpu_model = "not a file"
            cpu_core = "not a file"
            cpu_thread = "not a file"
            cpu_count = "not a file"

            memory_sn = "not a file"
            memory_vendor = "not a file"
            memory_model = "not a file"
            memory_type = "not a file"
            memory_maximum_speed = "not a file"
            memory_configured_speed = "not a file"
            memory_capacity = "not a file"
            memory_count = "not a file"

            psu_sn = "not a file"
            psu_capacity = "not a file"
            psu_count = "not a file"

            ocp1_model = "not a file"
            ocp1_vendor = "not a file"
            ocp1_mac = "not a file"
            ocp1_slot = "not a file"

            pn = "not a file"
            bios_version = "not a file"
            bmc_version = "not a file"
            mb_cpld_version = "not a file"
            bp1_cpld_version = "not a file"
            bp2_cpld_version = "not a file"

            bios_mismatch_list = "not a file"
            bios_mismatch_count = "not a file"
            bios_configuration_status = "not a file"
            bios_ocp1_status = "not a file"
            bios_disable_boot_drive_status = "not a file"
        
        # 추출한 데이터를 데이터프레임에 추가
        df_data.append({
            "BMC IP": bmc_ip,
            "PN": pn,
            "SN": sn,
            "CPU SN": cpu_sn,
            "CPU Vendor": cpu_vendor,
            "CPU Model": cpu_model,
            "CPU Core": cpu_core,
            "CPU Thread": cpu_thread,
            "CPU Count": cpu_count,
            "Memory SN": memory_sn,
            "Memory Vendor": memory_vendor,
            "Memory Model": memory_model,
            "Memory Type": memory_type,
            "Memory Maximum Speed": memory_maximum_speed,
            "Memory Configured Speed": memory_configured_speed,
            "Memory Capacity": memory_capacity,
            "Memory Count": memory_count,
            "NVMe SN": nvme_sn,
            "NVMe Vendor": nvme_vendor,
            "NVMe Model": nvme_model,
            "NVMe Interface": nvme_interface,
            "NVMe Capacity": nvme_capacity,
            "NVMe FW": nvme_fw,
            "NVMe Count": nvme_count,
            "OCP1 Model": ocp1_model,
            "OCP1 Vendor": ocp1_vendor,
            "OCP1 MAC": ocp1_mac,
            "OCP1 Slot": ocp1_slot,
            "PSU SN": psu_sn,
            "PSU Capacity": psu_capacity,
            "PSU Count": psu_count,
            "BIOS Version": bios_version,
            "BMC Version": bmc_version,
            "MB CPLD Version": mb_cpld_version,
            "BP1 CPLD Version": bp1_cpld_version,
            "BP2 CPLD Version": bp2_cpld_version,
            "BIOS Mismatch Count": bios_mismatch_count,
            "BIOS Configuration Status": bios_configuration_status,
            "BIOS Mismatch List": bios_mismatch_list,
            "BIOS OCP1 Status": bios_ocp1_status,
            "BIOS Disable Boot Drive Status": bios_disable_boot_drive_status
        })

    # 데이터프레임 생성 및 엑셀로 출력
    df = pd.DataFrame(data=df_data)
    df["ip sort"] = df["BMC IP"].apply(lambda x: int(ipaddress.IPv4Address(x)))
    df = df.sort_values(by="ip sort").drop(columns=["ip sort"])
    df.set_index("BMC IP", inplace=True)
    df.to_excel(output_file)
 
    # 엑셀 파일의 열 너비 자동 조정
    wb = opx.load_workbook(output_file)
    ws = wb.active
    for column_cells in ws.columns:
        column = column_cells[0].column_letter
        longest_cell = ""
        for cell in column_cells:
            if cell.value is not None:
                longest_cell = cell.value if len(str(cell.value)) > len(str(longest_cell)) else longest_cell
        max_length = len(str(longest_cell)) + 3
        ws.column_dimensions[column].width = max_length
    wb.save(output_file)

if __name__ == "__main__":
    main()

# 티콘서버도 구분
# 모든 카카오 서버도 구분
# nic가 2개인 경우에 대한 대처