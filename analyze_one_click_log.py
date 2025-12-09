from pathlib import Path
import tarfile as tf
import openpyxl as opx
import pandas as pd
import re
import ipaddress

# 사용자가 변경해야할 값들
folder = Path(r"C:\Users\lionm\Downloads\20250415_납품_SC_100_김정우-20251209T153450Z-1-001\20250415_납품_SC_100_김정우\2차 원클릭 로그\LogCollectResult")
output_file = Path(r"C:\Users\lionm\Downloads\output_one_click_log.xlsx")

bios_config_dict_1258h_v7 = {
    "EnableSVM": "Disabled",
    "OptionValue3": "Disabled",
    "CbsCmnGnbNbIOMMU": "Auto",
    "OptionValue1": "Disabled",
    "CbsCpuSmtCtrl": "Enabled",
    "CbsDfCmnDramNps": "NPS1",
    "CbsDfCmnMemIntlv": "Enabled",
    "CbsDfCmnAcpiSratL3Numa": "Disabled",
    "CbsCmnCpuCpb": "Disabled",
    "CbsCmnCpuGlobalCstateCtrl": "Disabled",
    "CbsCmnCpuMonMwaitDis": "Disabled",
    "CbsCmnApbdis": "1",
    "CbsCmnApbdisDfPstateRs": "0",
    "CbsCmnGnbSmuDfCstatesRs": "Disabled",
    "CbsCmnNbioPcieIdlePowerSetting": "Optimize for Perf/Power",
    "BootType": "UEFIBoot",
    "HttpNetworkProtocol": "Disabled",
    "UsbBoot": "Disabled",
    "SyncOrder": "Disabled",
    "BootTypeOrder0": "HardDiskDrive",
    "BootTypeOrder1": "PXE",
    "BootTypeOrder2": "DVDROMDrive",
    "BootTypeOrder3": "Others"
}

bios_config_dict_k1_k2_k3_k4_sa_2288h_v7_l4 = {
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

bios_config_dict_k5 = {
    "PcieAspmControl": "Disabled",
    "SriovEnablePolicy": "Enabled",
    "ProcessorHyperThreadingDisable": "ALL LPs",
    "ProcessorVmxEnable": "Enabled",
    "VTdSupport": "Enabled",
    "ProcessorX2apic": "Enabled",
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

bios_config_dict_sb_sc = {
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
    "PciePortEnable9": "Only Disable Boot Driver",
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

                # app_revision.txt >> pn 정보 추출
                app_revision_txt_tarinfo = tar.getmember("dump_info/RTOSDump/versioninfo/app_revision.txt")
                if app_revision_txt_tarinfo.isfile():
                    with tar.extractfile(app_revision_txt_tarinfo) as app_revision_txt_file:
                        app_revision_txt_content = app_revision_txt_file.read().decode('utf-8')
                        pn_regex = re.search(r'Product\s*Name:\s*([^\n\r]+)', app_revision_txt_content)
                        if pn_regex:
                            pn = pn_regex.group(1)
                        else:
                            print("PN not found in app_revision.txt for file")
                            exit()

                match pn:
                    case "2258 V7":
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

                        # server_config.txt >> cpu, memory, nvme, ssd, hdd, raid, psu, gpu, ocp 정보 추출
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

                                    cpu_vendor = [line.split()[4] for line in valid_cpu_info]
                                    cpu_model = [" ".join(line.split()[5:8]).rstrip(',') for line in valid_cpu_info]
                                    cpu_core = [line.split(',')[4].lstrip() for line in valid_cpu_info]
                                    cpu_thread = [line.split(',')[5].lstrip() for line in valid_cpu_info]
                                    cpu_count = len(valid_cpu_info)

                                # cpu info 섹션 자체를 찾지 못한 경우
                                else:
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
                                    memory_type = [line.split(',')[7].lstrip() for line in valid_memory_info]
                                    memory_maximum_speed = [line.split(',')[5].lstrip() for line in valid_memory_info]
                                    memory_configured_speed = [line.split(',')[6].lstrip() for line in valid_memory_info]
                                    memory_capacity = [line.split(',')[4].lstrip() for line in valid_memory_info]
                                    memory_slot = [line.split(',')[2].lstrip() for line in valid_memory_info]
                                    memory_count = len(valid_memory_info)
                                else:
                                    memory_sn = "memory info not found in server_config.txt"
                                    memory_vendor = "memory info not found in server_config.txt"
                                    memory_type = "memory info not found in server_config.txt"
                                    memory_maximum_speed = "memory info not found in server_config.txt"
                                    memory_configured_speed = "memory info not found in server_config.txt"
                                    memory_capacity = "memory info not found in server_config.txt"
                                    memory_slot = "memory info not found in server_config.txt"
                                    memory_count = "memory info not found in server_config.txt"

                                # RAID Card1 정보 추출
                                raid_card1_info1_regex = re.search(r"RAID Card Info.*?(?=\n\s*\n)", server_config_txt_content, re.DOTALL)
                                if raid_card1_info1_regex:
                                    raid_card1_info1 = raid_card1_info1_regex.group(0)

                                    raid_card1_model = [line.split("|")[2].strip() for line in raid_card1_info1.splitlines()[2:]]
                                    raid_card1_vendor = [line.split("|")[3].strip() for line in raid_card1_info1.splitlines()[2:]]
                                else:
                                    raid_card1_model = "raid card1 info1 not found in server_config.txt"
                                    raid_card1_vendor = "raid card1 info1 not found in server_config.txt"

                                raid_card1_info2_regex = re.search(r"(RAID Controller #0 Information.*)(?=\n\s*\nPass Through Drives Information)", server_config_txt_content, re.DOTALL)
                                if raid_card1_info2_regex:
                                    raid_card1_info2 = raid_card1_info2_regex.group(1)
                                    raid_card1_cache_regex = re.search(r"Memory Size\s*:\s*([^\n\r]+)", raid_card1_info2)
                                    raid_card1_cache = raid_card1_cache_regex.group(1).lstrip() if raid_card1_cache_regex else "raid card1 cache size not found in server_config.txt"
                                    raid_card1_fw_regex = re.search(r"Firmware Version\s*:\s*([^\n\r]+)", raid_card1_info2)
                                    raid_card1_fw = raid_card1_fw_regex.group(1).lstrip() if raid_card1_fw_regex else "raid card1 fw version not found in server_config.txt"
                                    raid_card1_nvdata_regex = re.search(r"NVDATA Version\s*:\s*([^\n\r]+)", raid_card1_info2)
                                    raid_card1_nvdata = raid_card1_nvdata_regex.group(1).lstrip() if raid_card1_nvdata_regex else "raid card1 nvdata version not found in server_config.txt"
                                    raid_card1_level_regex = re.search(r"^Type\s*:\s*([^\n\r]+)", raid_card1_info2, re.MULTILINE)
                                    raid_card1_level = raid_card1_level_regex.group(1).lstrip() if raid_card1_level_regex else "raid card1 raid level not found in server_config.txt"
                                    raid_card1_capacity_regex = re.search(r"^Total Size\s*:\s*([^\n\r]+)", raid_card1_info2, re.MULTILINE)
                                    raid_card1_capacity = raid_card1_capacity_regex.group(1).lstrip() if raid_card1_capacity_regex else "raid card1 capacity not found in server_config.txt"
                                else:
                                    raid_card1_cache = "raid card1 info2 not found in server_config.txt"
                                    raid_card1_fw = "raid card1 info2 not found in server_config.txt"
                                    raid_card1_nvdata = "raid card1 info2 not found in server_config.txt"
                                    raid_card1_level = "raid card1 info2 not found in server_config.txt"
                                    raid_card1_capacity = "raid card1 info2 not found in server_config.txt" 

                                # RAID Card1 Disk 정보 추출
                                raid_card1_disk_info_regex = re.search(r"(Physical Drives Information.*)", raid_card1_info2, re.DOTALL)
                                if raid_card1_disk_info_regex:
                                    raid_card1_disk_info = raid_card1_disk_info_regex.group(1)
                                    raid_card1_disk_sn_list = re.findall(r"Serial Number\s*:\s*[^\n\r]+", raid_card1_disk_info)
                                    raid_card1_disk_sn = [line.split(":")[1].lstrip().rstrip() for line in raid_card1_disk_sn_list]
                                    raid_card1_disk_vendor_list = re.findall(r"Manufacturer\s*:\s*[^\n\r]+", raid_card1_disk_info)
                                    raid_card1_disk_vendor = [line.split(":")[1].lstrip().rstrip() for line in raid_card1_disk_vendor_list]
                                    raid_card1_disk_model_list = re.findall(r"Model\s*:\s*[^\n\r]+", raid_card1_disk_info)
                                    raid_card1_disk_model = [line.split()[-1].lstrip().rstrip() for line in raid_card1_disk_model_list]
                                    raid_card1_disk_capacity_list = re.findall(r"Capacity\s*:\s*[^\n\r]+", raid_card1_disk_info)
                                    raid_card1_disk_capacity = [line.split(":")[1].lstrip().rstrip() for line in raid_card1_disk_capacity_list]
                                    raid_card1_disk_count = len(re.findall(r"^ID\s*:\s*[^\n\r]+", raid_card1_disk_info, re.MULTILINE))
                                else:
                                    raid_card1_disk_sn = "raid card1 disk info not found in server_config.txt"
                                    raid_card1_disk_vendor = "raid card1 disk info not found in server_config.txt"
                                    raid_card1_disk_model = "raid card1 disk info not found in server_config.txt"
                                    raid_card1_disk_capacity = "raid card1 disk info not found in server_config.txt"
                                    raid_card1_disk_count = "raid card1 disk info not found in server_config.txt"

                                # GPU 정보 추출
                                gpu_info_regex = re.search(r"GPU Card Info.*?(?=\n\s*\n)", server_config_txt_content, re.DOTALL)
                                if gpu_info_regex:
                                    gpu_info = gpu_info_regex.group(0)
                                    gpu_sn = [line.split("|")[5].strip() for line in gpu_info.splitlines()[2:]]
                                    gpu_model = [line.split("|")[3].strip() for line in gpu_info.splitlines()[2:]]
                                    gpu_slot = [line.split("|")[0].strip() for line in gpu_info.splitlines()[2:]]
                                else:
                                    gpu_sn = "gpu info not found in server_config.txt"
                                    gpu_model = "gpu info not found in server_config.txt"
                                    gpu_slot = "gpu info not found in server_config.txt"

                                # PSU 정보 추출
                                psu_info_regex = re.search(r"(-+PSU\s*info.*)(?=\n\s*\n-*Storage\s*info)", server_config_txt_content, re.DOTALL)
                                if psu_info_regex:
                                    psu_info = psu_info_regex.group(1)
                                    valid_psu_info = [line for line in psu_info.splitlines()[2:]]
                                    psu_capacity = [line.split('|')[6].lstrip().rstrip() for line in valid_psu_info]
                                    psu_count = len(valid_psu_info)

                                # psu info 섹션 자체를 찾지 못한 경우
                                else:
                                    psu_capacity = "psu info not found in server_config.txt"
                                    psu_count = "psu info not found in server_config.txt"

                        # server_config.txt 파일 자체를 찾지 못한 경우
                        else:
                            cpu_vendor = "no server_config.txt file"
                            cpu_model = "no server_config.txt file"
                            cpu_core = "no server_config.txt file"
                            cpu_thread = "no server_config.txt file"
                            cpu_count = "no server_config.txt file"

                            memory_sn = "no server_config.txt file"
                            memory_vendor = "no server_config.txt file"
                            memory_type = "no server_config.txt file"
                            memory_maximum_speed = "no server_config.txt file"
                            memory_configured_speed = "no server_config.txt file"
                            memory_capacity = "no server_config.txt file"
                            memory_slot = "no server_config.txt file"
                            memory_count = "no server_config.txt file"

                            raid_card1_disk_sn = "no server_config.txt file"
                            raid_card1_disk_vendor = "no server_config.txt file"
                            raid_card1_disk_model = "no server_config.txt file"
                            raid_card1_disk_capacity = "no server_config.txt file"
                            raid_card1_disk_count = "no server_config.txt file"

                            raid_card1_model = "no server_config.txt file"
                            raid_card1_vendor = "no server_config.txt file"
                            raid_card1_cache = "no server_config.txt file"
                            raid_card1_fw = "no server_config.txt file"
                            raid_card1_nvdata = "no server_config.txt file"
                            raid_card1_level = "no server_config.txt file"
                            raid_card1_capacity = "no server_config.txt file"

                            gpu_sn = "no server_config.txt file"
                            gpu_model = "no server_config.txt file"
                            gpu_slot = "no server_config.txt file"

                            psu_capacity = "no server_config.txt file"
                            psu_count = "no server_config.txt file"

                        # netcard_info.txt >> pcie nic 정보 추출
                        netcard_info_tarinfo = tar.getmember("dump_info/LogDump/netcard/netcard_info.txt")
                        if netcard_info_tarinfo.isfile():
                            with tar.extractfile(netcard_info_tarinfo) as netcard_info_file:
                                netcard_info_content = netcard_info_file.read().decode('utf-8')

                                utc_match_list = list(re.finditer(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} UTC', netcard_info_content))
                                if utc_match_list:
                                    nic_info2 = netcard_info_content[utc_match_list[-1].start():]

                                    idx = next(i for i, match in enumerate(nic_info2.splitlines()) if "SlotId          :1" in match)
                                    nic1_info2 = "\n".join(nic_info2.splitlines()[idx-3:idx+6])
                                    nic1_model_regex = re.search(r'ProductName\s*:([^\n\r]+)', nic1_info2)
                                    nic1_model = nic1_model_regex.group(1).lstrip() if nic1_model_regex else "not found in netcard_info.txt"
                                    nic1_mac_regex = re.search(r'MacAddr\s*:([^\n\r]+)', nic1_info2)
                                    nic1_mac = nic1_mac_regex.group(1).lstrip() if nic1_mac_regex else "not found in netcard_info.txt"
                                    nic1_slot_regex = re.search(r'SlotId\s*:([^\n\r]+)', nic1_info2)
                                    nic1_slot = nic1_slot_regex.group(1).lstrip() if nic1_slot_regex else "not found in netcard_info.txt"
                                    nic1_vendor_regex = re.search(r'Manufacture\s*:([^\n\r]+)', nic1_info2)
                                    nic1_vendor = nic1_vendor_regex.group(1).lstrip() if nic1_vendor_regex else "not found in netcard_info.txt"

                                    idx = next(i for i, match in enumerate(nic_info2.splitlines()) if "SlotId          :2" in match)
                                    nic2_info2 = "\n".join(nic_info2.splitlines()[idx-3:idx+6])
                                    nic2_model_regex = re.search(r'ProductName\s*:([^\n\r]+)', nic2_info2)
                                    nic2_model = nic2_model_regex.group(1).lstrip() if nic2_model_regex else "not found in netcard_info.txt"
                                    nic2_mac_regex = re.search(r'MacAddr\s*:([^\n\r]+)', nic2_info2)
                                    nic2_mac = nic2_mac_regex.group(1).lstrip() if nic2_mac_regex else "not found in netcard_info.txt"
                                    nic2_slot_regex = re.search(r'SlotId\s*:([^\n\r]+)', nic2_info2)
                                    nic2_slot = nic2_slot_regex.group(1).lstrip() if nic2_slot_regex else "not found in netcard_info.txt"
                                    nic2_vendor_regex = re.search(r'Manufacture\s*:([^\n\r]+)', nic2_info2)
                                    nic2_vendor = nic2_vendor_regex.group(1).lstrip() if nic2_vendor_regex else "not found in netcard_info.txt"
                                else:
                                    nic1_model = "not found in netcard_info.txt"
                                    nic1_mac = "not found in netcard_info.txt"
                                    nic1_slot = "not found in netcard_info.txt"
                                    nic1_vendor = "not found in netcard_info.txt"
                                    nic2_model = "not found in netcard_info.txt"
                                    nic2_mac = "not found in netcard_info.txt"
                                    nic2_slot = "not found in netcard_info.txt"
                                    nic2_vendor = "not found in netcard_info.txt"

                        # netcard_info.txt 파일 자체를 찾지 못한 경우
                        else:
                            nic1_model = "no netcard_info.txt file"
                            nic1_mac = "no netcard_info.txt file"
                            nic1_slot = "no netcard_info.txt file"
                            nic1_vendor = "no netcard_info.txt file"
                            nic2_model = "no netcard_info.txt file"
                            nic2_mac = "no netcard_info.txt file"
                            nic2_slot = "no netcard_info.txt file"
                            nic2_vendor = "no netcard_info.txt file"

                        # app_revision.txt >> bios, bmc 정보 추출
                        app_revision_txt_tarinfo = tar.getmember("dump_info/RTOSDump/versioninfo/app_revision.txt")
                        if app_revision_txt_tarinfo.isfile():
                            with tar.extractfile(app_revision_txt_tarinfo) as app_revision_txt_file:
                                app_revision_txt_content = app_revision_txt_file.read().decode('utf-8')
                                bios_version_regex = re.search(r'Active BIOS\s+Version\s*:\s*([^\n\r]+)', app_revision_txt_content)
                                bios_version = bios_version_regex.group(1) if bios_version_regex else "not found in app_revision.txt"
                                bmc_version_regex = re.search(r'Active iBMC\s+Version\s*:\s*([^\n\r]+)', app_revision_txt_content)
                                bmc_version = bmc_version_regex.group(1) if bmc_version_regex else "not found in app_revision.txt"

                        # app_revision.txt 파일 자체를 찾지 못한 경우
                        else:
                            bios_version = "no app_revision.txt file"
                            bmc_version = "no app_revision.txt file"

                        # 추출한 데이터를 데이터프레임에 추가
                        df_data.append({
                            "BMC IP": bmc_ip,
                            "PN": pn,
                            "SN": sn,
                            "CPU Vendor": cpu_vendor,
                            "CPU Model": cpu_model,
                            "CPU Core": cpu_core,
                            "CPU Thread": cpu_thread,
                            "CPU Count": cpu_count,
                            "Memory SN": memory_sn,
                            "Memory Vendor": memory_vendor,
                            "Memory Type": memory_type,
                            "Memory Maximum Speed": memory_maximum_speed,
                            "Memory Configured Speed": memory_configured_speed,
                            "Memory Capacity": memory_capacity,
                            "Memory Slot": memory_slot,
                            "Memory Count": memory_count,
                            "RAID Card1 Disk SN": raid_card1_disk_sn,
                            "RAID Card1 Disk Vendor": raid_card1_disk_vendor,
                            "RAID Card1 Disk Model": raid_card1_disk_model,
                            "RAID Card1 Disk Capacity": raid_card1_disk_capacity,
                            "RAID Card1 Disk Count": raid_card1_disk_count,
                            "RAID Card1 Model": raid_card1_model,
                            "RAID Card1 Vendor": raid_card1_vendor,
                            "RAID Card1 Cache": raid_card1_cache,
                            "RAID Card1 FW": raid_card1_fw,
                            "RAID Card1 NVDATA": raid_card1_nvdata,
                            "RAID Card1 Level": raid_card1_level,
                            "RAID Card1 Capacity": raid_card1_capacity,
                            "NIC1 Model": nic1_model,
                            "NIC1 MAC": nic1_mac,
                            "NIC1 SLOT": nic1_slot,
                            "NIC1 VENDOR": nic1_vendor,
                            "NIC2 Model": nic2_model,
                            "NIC2 MAC": nic2_mac,
                            "NIC2 SLOT": nic2_slot,
                            "NIC2 VENDOR": nic2_vendor,
                            "PSU Capacity": psu_capacity,
                            "PSU Count": psu_count,
                            "GPU SN": gpu_sn,
                            "GPU Model": gpu_model,
                            "GPU Slot": gpu_slot,
                            "BIOS Version": bios_version,
                            "BMC Version": bmc_version
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
                    case _:
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

                        # server_config.txt >> cpu, memory, nvme, ssd, hdd, raid, psu, gpu, ocp 정보 추출
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

                                # RAID Card1 정보 추출
                                raid_card1_info1_regex = re.search(r"RAID Card Info.*?(?=\n\s*\n)", server_config_txt_content, re.DOTALL)
                                if raid_card1_info1_regex:
                                    raid_card1_info1 = raid_card1_info1_regex.group(0)

                                    raid_card1_model = [line.split("|")[2].strip() for line in raid_card1_info1.splitlines()[2:]]
                                    raid_card1_vendor = [line.split("|")[3].strip() for line in raid_card1_info1.splitlines()[2:]]
                                else:
                                    raid_card1_model = "raid card1 info1 not found in server_config.txt"
                                    raid_card1_vendor = "raid card1 info1 not found in server_config.txt"

                                raid_card1_info2_regex = re.search(r"(RAID Controller #0 Information.*)(?=\n\s*\nRAID Controller #1 Information)", server_config_txt_content, re.DOTALL)
                                if raid_card1_info2_regex:
                                    raid_card1_info2 = raid_card1_info2_regex.group(1)
                                    raid_card1_cache_regex = re.search(r"Memory Size\s*:\s*([^\n\r]+)", raid_card1_info2)
                                    raid_card1_cache = raid_card1_cache_regex.group(1).lstrip() if raid_card1_cache_regex else "raid card1 cache size not found in server_config.txt"
                                    raid_card1_fw_regex = re.search(r"Firmware Version\s*:\s*([^\n\r]+)", raid_card1_info2)
                                    raid_card1_fw = raid_card1_fw_regex.group(1).lstrip() if raid_card1_fw_regex else "raid card1 fw version not found in server_config.txt"
                                    raid_card1_nvdata_regex = re.search(r"NVDATA Version\s*:\s*([^\n\r]+)", raid_card1_info2)
                                    raid_card1_nvdata = raid_card1_nvdata_regex.group(1).lstrip() if raid_card1_nvdata_regex else "raid card1 nvdata version not found in server_config.txt"
                                    raid_card1_level_regex = re.search(r"^Type\s*:\s*([^\n\r]+)", raid_card1_info2, re.MULTILINE)
                                    raid_card1_level = raid_card1_level_regex.group(1).lstrip() if raid_card1_level_regex else "raid card1 raid level not found in server_config.txt"
                                    raid_card1_capacity_regex = re.search(r"^Total Size\s*:\s*([^\n\r]+)", raid_card1_info2, re.MULTILINE)
                                    raid_card1_capacity = raid_card1_capacity_regex.group(1).lstrip() if raid_card1_capacity_regex else "raid card1 capacity not found in server_config.txt"

                                else:
                                    raid_card1_cache = "raid card1 info2 not found in server_config.txt"
                                    raid_card1_fw = "raid card1 info2 not found in server_config.txt"
                                    raid_card1_nvdata = "raid card1 info2 not found in server_config.txt"
                                    raid_card1_level = "raid card1 info2 not found in server_config.txt"
                                    raid_card1_capacity = "raid card1 info2 not found in server_config.txt" 

                                # RAID Card1 Disk 정보 추출
                                raid_card1_disk_info_regex = re.search(r"(Physical Drives Information.*)", raid_card1_info2, re.DOTALL)
                                if raid_card1_info2_regex and raid_card1_disk_info_regex:
                                    raid_card1_disk_info = raid_card1_disk_info_regex.group(1)
                                    raid_card1_disk_sn_list = re.findall(r"Serial Number\s*:\s*[^\n\r]+", raid_card1_disk_info)
                                    raid_card1_disk_sn = [line.split(":")[1].lstrip().rstrip() for line in raid_card1_disk_sn_list]
                                    raid_card1_disk_vendor_list = re.findall(r"Manufacturer\s*:\s*[^\n\r]+", raid_card1_disk_info)
                                    raid_card1_disk_vendor = [line.split(":")[1].lstrip().rstrip() for line in raid_card1_disk_vendor_list]
                                    raid_card1_disk_model_list = re.findall(r"Model\s*:\s*[^\n\r]+", raid_card1_disk_info)
                                    raid_card1_disk_model = [line.split()[-1].lstrip().rstrip() for line in raid_card1_disk_model_list]
                                    raid_card1_disk_interface_list = re.findall(r"Interface Type\s*:\s*[^\n\r]+", raid_card1_disk_info)
                                    raid_card1_disk_interface = [line.split(":")[1].lstrip().rstrip() for line in raid_card1_disk_interface_list]
                                    raid_card1_disk_capacity_list = re.findall(r"Capacity\s*:\s*[^\n\r]+", raid_card1_disk_info)
                                    raid_card1_disk_capacity = [line.split(":")[1].lstrip().rstrip() for line in raid_card1_disk_capacity_list]
                                    raid_card1_disk_fw_list = re.findall(r"Firmware Version\s*:\s*[^\n\r]+", raid_card1_disk_info)
                                    raid_card1_disk_fw = [line.split(":")[1].lstrip().rstrip() for line in raid_card1_disk_fw_list]
                                    raid_card1_disk_count = len(re.findall(r"^ID\s*:\s*[^\n\r]+", raid_card1_disk_info, re.MULTILINE))
                                else:
                                    raid_card1_disk_sn = "raid card1 disk info not found in server_config.txt"
                                    raid_card1_disk_vendor = "raid card1 disk info not found in server_config.txt"
                                    raid_card1_disk_model = "raid card1 disk info not found in server_config.txt"
                                    raid_card1_disk_interface = "raid card1 disk info not found in server_config.txt"
                                    raid_card1_disk_capacity = "raid card1 disk info not found in server_config.txt"
                                    raid_card1_disk_fw = "raid card1 disk info not found in server_config.txt"
                                    raid_card1_disk_count = "raid card1 disk info not found in server_config.txt"

                                # RAID Card2 정보 추출
                                raid_card2_info1_regex = re.search(r"Pcie Card Info.*?(?=\n\s*\n)", server_config_txt_content, re.DOTALL)
                                if raid_card2_info1_regex:
                                    raid_card2_info1 = raid_card2_info1_regex.group(0)

                                    raid_card2_model = [line.split("|")[9].strip() for line in raid_card2_info1.splitlines()[2:]]
                                else:
                                    raid_card2_model = "raid card2 info1 not found in server_config.txt"

                                raid_card2_info2_regex = re.search(r"(RAID Controller #1 Information.*)(?=\n\s*\nPass Through Drives Information)", server_config_txt_content, re.DOTALL)
                                if raid_card2_info2_regex:
                                    raid_card2_info2 = raid_card2_info2_regex.group(1)
                                    raid_card2_cache_regex = re.search(r"Memory Size\s*:\s*([^\n\r]+)", raid_card2_info2)
                                    raid_card2_cache = raid_card2_cache_regex.group(1).lstrip() if raid_card2_cache_regex else "raid card2 cache size not found in server_config.txt"
                                    raid_card2_fw_regex = re.search(r"Firmware Version\s*:\s*([^\n\r]+)", raid_card2_info2)
                                    raid_card2_fw = raid_card2_fw_regex.group(1).lstrip() if raid_card2_fw_regex else "raid card2 fw version not found in server_config.txt"
                                    raid_card2_nvdata_regex = re.search(r"NVDATA Version\s*:\s*([^\n\r]+)", raid_card2_info2)
                                    raid_card2_nvdata = raid_card2_nvdata_regex.group(1).lstrip() if raid_card2_nvdata_regex else "raid card2 nvdata version not found in server_config.txt"
                                    raid_card2_jbod_regex = re.search(r"^JBOD Enabled\s*:\s*([^\n\r]+)", raid_card2_info2, re.MULTILINE)
                                    raid_card2_jbod = raid_card2_jbod_regex.group(1).lstrip() if raid_card2_jbod_regex else "raid card2 jbod status not found in server_config.txt"

                                else:
                                    raid_card2_cache = "raid card2 info2 not found in server_config.txt"
                                    raid_card2_fw = "raid card2 info2 not found in server_config.txt"
                                    raid_card2_nvdata = "raid card2 info2 not found in server_config.txt"
                                    raid_card2_jbod = "raid card2 info2 not found in server_config.txt"

                                # RAID Card2 Disk 정보 추출
                                raid_card2_disk_info_regex = re.search(r"(Physical Drives Information.*)", raid_card2_info2, re.DOTALL)
                                if raid_card2_info2_regex and raid_card2_disk_info_regex:
                                    raid_card2_disk_info = raid_card2_disk_info_regex.group(1)
                                    raid_card2_disk_sn_list = re.findall(r"Serial Number\s*:\s*[^\n\r]+", raid_card2_disk_info)
                                    raid_card2_disk_sn = [line.split(":")[1].lstrip().rstrip() for line in raid_card2_disk_sn_list]
                                    raid_card2_disk_vendor_list = re.findall(r"Manufacturer\s*:\s*[^\n\r]+", raid_card2_disk_info)
                                    raid_card2_disk_vendor = [line.split(":")[1].lstrip().rstrip() for line in raid_card2_disk_vendor_list]
                                    raid_card2_disk_model_list = re.findall(r"Model\s*:\s*[^\n\r]+", raid_card2_disk_info)
                                    raid_card2_disk_model = [line.split()[-1].lstrip().rstrip() for line in raid_card2_disk_model_list]
                                    raid_card2_disk_interface_list = re.findall(r"Interface Type\s*:\s*[^\n\r]+", raid_card2_disk_info)
                                    raid_card2_disk_interface = [line.split(":")[1].lstrip().rstrip() for line in raid_card2_disk_interface_list]
                                    raid_card2_disk_capacity_list = re.findall(r"Capacity\s*:\s*[^\n\r]+", raid_card2_disk_info)
                                    raid_card2_disk_capacity = [line.split(":")[1].lstrip().rstrip() for line in raid_card2_disk_capacity_list]
                                    raid_card2_disk_fw_list = re.findall(r"Firmware Version\s*:\s*[^\n\r]+", raid_card2_disk_info)
                                    raid_card2_disk_fw = [line.split(":")[1].lstrip().rstrip() for line in raid_card2_disk_fw_list]
                                    raid_card2_disk_simplified_fw_list = list(set(raid_card2_disk_fw))
                                    raid_card2_disk_count = len(re.findall(r"^ID\s*:\s*[^\n\r]+", raid_card2_disk_info, re.MULTILINE))
                                else:
                                    raid_card2_disk_sn = "raid card2 disk info not found in server_config.txt"
                                    raid_card2_disk_vendor = "raid card2 disk info not found in server_config.txt"
                                    raid_card2_disk_model = "raid card2 disk info not found in server_config.txt"
                                    raid_card2_disk_interface = "raid card2 disk info not found in server_config.txt"
                                    raid_card2_disk_capacity = "raid card2 disk info not found in server_config.txt"
                                    raid_card2_disk_fw = "raid card2 disk info not found in server_config.txt"
                                    raid_card2_disk_count = "raid card2 disk info not found in server_config.txt"

                                # GPU 정보 추출
                                gpu_info_regex = re.search(r"GPU Card Info.*?(?=\n\s*\n)", server_config_txt_content, re.DOTALL)
                                if gpu_info_regex:
                                    gpu_info = gpu_info_regex.group(0)
                                    gpu_sn = [line.split("|")[5].strip() for line in gpu_info.splitlines()[2:]]
                                    gpu_model = [line.split("|")[3].strip() for line in gpu_info.splitlines()[2:]]
                                    gpu_fw = [line.split("|")[-1].strip() for line in gpu_info.splitlines()[2:]]
                                    gpu_count = len(gpu_info.splitlines()[2:])
                                else:
                                    gpu_sn = "gpu info not found in server_config.txt"
                                    gpu_model = "gpu info not found in server_config.txt"
                                    gpu_fw = "gpu info not found in server_config.txt"
                                    gpu_count = "gpu info not found in server_config.txt"

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

                            nvme_sn = "no server_config.txt file"
                            nvme_vendor = "no server_config.txt file"
                            nvme_model = "no server_config.txt file"
                            nvme_interface = "no server_config.txt file"
                            nvme_capacity = "no server_config.txt file"
                            nvme_fw = "no server_config.txt file"
                            nvme_count = "no server_config.txt file"
                            
                            raid_card1_model = "no server_config.txt file"
                            raid_card1_vendor = "no server_config.txt file"
                            raid_card1_cache = "no server_config.txt file"
                            raid_card1_fw = "no server_config.txt file"
                            raid_card1_nvdata = "no server_config.txt file"
                            raid_card1_level = "no server_config.txt file"
                            raid_card1_capacity = "no server_config.txt file"

                            raid_card1_disk_sn = "no server_config.txt file"
                            raid_card1_disk_vendor = "no server_config.txt file"
                            raid_card1_disk_model = "no server_config.txt file" 
                            raid_card1_disk_interface = "no server_config.txt file"
                            raid_card1_disk_capacity = "no server_config.txt file"
                            raid_card1_disk_fw = "no server_config.txt file"
                            raid_card1_disk_count = "no server_config.txt file"

                            raid_card2_model = "no server_config.txt file"
                            raid_card2_cache = "no server_config.txt file"
                            raid_card2_fw = "no server_config.txt file"
                            raid_card2_nvdata = "no server_config.txt file"
                            raid_card2_jbod = "no server_config.txt file"

                            raid_card2_disk_sn = "no server_config.txt file"
                            raid_card2_disk_vendor = "no server_config.txt file"
                            raid_card2_disk_model = "no server_config.txt file"
                            raid_card2_disk_interface = "no server_config.txt file"
                            raid_card2_disk_capacity = "no server_config.txt file"
                            raid_card2_disk_fw = "no server_config.txt file"
                            raid_card2_disk_count = "no server_config.txt file"

                            gpu_sn = "no server_config.txt file"
                            gpu_model = "no server_config.txt file"
                            gpu_fw = "no server_config.txt file"
                            gpu_count = "no server_config.txt file"

                            psu_sn = "no server_config.txt file"
                            psu_capacity = "no server_config.txt file"
                            psu_count = "no server_config.txt file"

                        # netcard_info.txt >> ocp 정보 추출
                        server_config_txt_tarinfo = tar.getmember("dump_info/RTOSDump/versioninfo/server_config.txt")
                        if server_config_txt_tarinfo.isfile():
                            with tar.extractfile(server_config_txt_tarinfo) as server_config_txt_file:
                                server_config_txt_content = server_config_txt_file.read().decode('utf-8')
                                ocp_info1_regex = re.search(r"OCP Card Info.*?(?=\n\s*\n)", server_config_txt_content, re.DOTALL)
                                if ocp_info1_regex:
                                    ocp_info1 = ocp_info1_regex.group(0)
                                    ocp1_info1 = [line for line in ocp_info1.splitlines() if line.lstrip().startswith("1")]
                                    ocp2_info1 = [line for line in ocp_info1.splitlines() if line.lstrip().startswith("2")]
                                    ocp1_sn = ocp1_info1[0].split("|")[-2].strip() if ocp1_info1 else "not found in server_config.txt"
                                    ocp2_sn = ocp2_info1[0].split("|")[-2].strip() if ocp2_info1 else "not found in server_config.txt"
                                else:
                                    ocp1_sn = "ocp info1 not found in server_config.txt"
                                    ocp2_sn = "ocp info1 not found in server_config.txt"
                        # server_config.txt 파일 자체를 찾지 못한 경우
                        else:
                            ocp1_sn = "no server_config.txt file"
                            ocp2_sn = "no server_config.txt file"

                        netcard_info_tarinfo = tar.getmember("dump_info/LogDump/netcard/netcard_info.txt")
                        if netcard_info_tarinfo.isfile():
                            with tar.extractfile(netcard_info_tarinfo) as netcard_info_file:
                                netcard_info_content = netcard_info_file.read().decode('utf-8')

                                utc_match_list = list(re.finditer(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} UTC', netcard_info_content))
                                if utc_match_list:
                                    ocp_info2 = netcard_info_content[utc_match_list[-1].start():]

                                    idx = next((i for i, match in enumerate(ocp_info2.splitlines()) if "SlotId          :1" in match), None)
                                    if idx is not None:
                                        ocp1_info2 = "\n".join(ocp_info2.splitlines()[idx-3:idx+6])
                                        ocp1_model_regex = re.search(r'ProductName\s*:([^\n\r]+)', ocp1_info2)
                                        ocp1_model = ocp1_model_regex.group(1).lstrip() if ocp1_model_regex else "not found in netcard_info.txt"
                                        ocp1_mac_regex = re.search(r'MacAddr\s*:([^\n\r]+)', ocp1_info2)
                                        ocp1_mac = ocp1_mac_regex.group(1).lstrip() if ocp1_mac_regex else "not found in netcard_info.txt"
                                        ocp1_fw_regex = re.search(r'FirmwareVersion\s*:([^\n\r]+)', ocp1_info2)
                                        ocp1_fw = ocp1_fw_regex.group(1).lstrip() if ocp1_fw_regex else "not found in netcard_info.txt"
                                    else:
                                        ocp1_model = "OCP1 not found in netcard_info.txt"
                                        ocp1_mac = "OCP1 not found in netcard_info.txt"
                                        ocp1_fw = "OCP1 not found in netcard_info.txt"

                                    idx = next((i for i, match in enumerate(ocp_info2.splitlines()) if "SlotId          :2" in match), None)
                                    if idx is not None:
                                        ocp2_info2 = "\n".join(ocp_info2.splitlines()[idx-3:idx+6])
                                        ocp2_model_regex = re.search(r'ProductName\s*:([^\n\r]+)', ocp2_info2)
                                        ocp2_model = ocp2_model_regex.group(1).lstrip() if ocp2_model_regex else "not found in netcard_info.txt"
                                        ocp2_mac_regex = re.search(r'MacAddr\s*:([^\n\r]+)', ocp2_info2)
                                        ocp2_mac = ocp2_mac_regex.group(1).lstrip() if ocp2_mac_regex else "not found in netcard_info.txt"
                                        ocp2_fw_regex = re.search(r'FirmwareVersion\s*:([^\n\r]+)', ocp2_info2)
                                        ocp2_fw = ocp2_fw_regex.group(1).lstrip() if ocp2_fw_regex else "not found in netcard_info.txt"
                                    else:
                                        ocp2_model = "OCP2 not found in netcard_info.txt"
                                        ocp2_mac = "OCP2 not found in netcard_info.txt"
                                        ocp2_fw = "OCP2 not found in netcard_info.txt"
                                else:
                                    ocp1_model = "not found in netcard_info.txt"
                                    ocp1_mac = "not found in netcard_info.txt"
                                    ocp1_fw = "not found in netcard_info.txt"
                                    ocp2_model = "not found in netcard_info.txt"
                                    ocp2_mac = "not found in netcard_info.txt"
                                    ocp2_fw = "not found in netcard_info.txt"

                        # netcard_info.txt 파일 자체를 찾지 못한 경우
                        else:
                            ocp1_model = "no netcard_info.txt file"
                            ocp1_mac = "no netcard_info.txt file"
                            ocp1_fw = "no netcard_info.txt file"
                            ocp2_model = "no netcard_info.txt file"
                            ocp2_mac = "no netcard_info.txt file"
                            ocp2_fw = "no netcard_info.txt file"

                        # app_revision.txt >> bios, bmc, cpld, vrd 정보 추출
                        app_revision_txt_tarinfo = tar.getmember("dump_info/RTOSDump/versioninfo/app_revision.txt")
                        if app_revision_txt_tarinfo.isfile():
                            with tar.extractfile(app_revision_txt_tarinfo) as app_revision_txt_file:
                                app_revision_txt_content = app_revision_txt_file.read().decode('utf-8')
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
                                vrd_version_regex = re.search(r'Mainboard\s*VRD\s*:\s*([^\n\r]+)', app_revision_txt_content)
                                vrd_version = vrd_version_regex.group(1) if vrd_version_regex else "not found in app_revision.txt"

                        # app_revision.txt 파일 자체를 찾지 못한 경우
                        else:
                            bios_version = "no app_revision.txt file"
                            bmc_version = "no app_revision.txt file"
                            mb_cpld_version = "no app_revision.txt file"
                            bp1_cpld_version = "no app_revision.txt file"
                            bp2_cpld_version = "no app_revision.txt file"
                            vrd_version = "no app_revision.txt file"

                        # BIOS 설정 정보 지정
                        #1258H V7 (AMD)
                        if pn == "1258HV7":
                            bios_config_dict = bios_config_dict_1258h_v7
                        #2288H V7 L4 (GPU)
                        elif pn == "2288HV7" and gpu_count > 0:
                            bios_config_dict = bios_config_dict_k1_k2_k3_k4_sa_2288h_v7_l4
                        #K5
                        elif cpu_model[0] == "Xeon(R) Gold 6430" and cpu_model[1] == "Xeon(R) Gold 6430":
                            bios_config_dict = bios_config_dict_k5
                        #SB
                        elif pn == "2288HV7" and gpu_count == 0:
                            bios_config_dict = bios_config_dict_sb_sc
                        #SC
                        elif pn == "5288V7":
                            bios_config_dict = bios_config_dict_sb_sc
                        #K1, K2, K3, K4, SA
                        else:
                            bios_config_dict = bios_config_dict_k1_k2_k3_k4_sa_2288h_v7_l4

                        # BIOS/currentvalue.json >> bios 설정 정보 추출
                        bios_currentvalue_json_tarinfo = tar.getmember("dump_info/AppDump/BIOS/currentvalue.json")
                        if bios_currentvalue_json_tarinfo.isfile():
                            with tar.extractfile(bios_currentvalue_json_tarinfo) as bios_currentvalue_json_file:
                                bios_currentvalue_json_content = bios_currentvalue_json_file.read().decode('utf-8')
                                bios_mismatch_list = []
                                bios_mismatch_count = 0
                                for key, expected_value in bios_config_dict.items():
                                    bios_actual_value_regex = re.search(rf'"{key}"\s*:\s*"([^"]+)"', bios_currentvalue_json_content)
                                    if bios_actual_value_regex:
                                        bios_actual_value = bios_actual_value_regex.group(1)
                                        if bios_actual_value != expected_value:
                                            bios_mismatch_list.append(key)
                                            bios_mismatch_count += 1
                                    else:
                                        bios_mismatch_list.append(key)
                                        bios_mismatch_count += 1
                                bios_configuration_status = "True" if bios_mismatch_count == 0 else "False"
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
                            "RAID Card1 Disk SN": raid_card1_disk_sn,
                            "RAID Card1 Disk Vendor": raid_card1_disk_vendor,
                            "RAID Card1 Disk Model": raid_card1_disk_model,
                            "RAID Card1 Disk Interface": raid_card1_disk_interface,
                            "RAID Card1 Disk Capacity": raid_card1_disk_capacity,
                            "RAID Card1 Disk FW": raid_card1_disk_fw,
                            "RAID Card1 Disk Count": raid_card1_disk_count,
                            "RAID Card2 Disk SN": raid_card2_disk_sn,
                            "RAID Card2 Disk Vendor": raid_card2_disk_vendor,
                            "RAID Card2 Disk Model": raid_card2_disk_model,
                            "RAID Card2 Disk Interface": raid_card2_disk_interface,
                            "RAID Card2 Disk Capacity": raid_card2_disk_capacity,
                            "RAID Card2 Disk FW": raid_card2_disk_fw,
                            "RAID Card2 Disk Simplified FW": raid_card2_disk_simplified_fw_list,
                            "RAID Card2 Disk Count": raid_card2_disk_count,
                            "RAID Card1 Vendor": raid_card1_vendor,
                            "RAID Card1 Model": raid_card1_model,
                            "RAID Card1 Cache": raid_card1_cache,
                            "RAID Card1 FW": raid_card1_fw,
                            "RAID Card1 NVDATA": raid_card1_nvdata,
                            "RAID Card1 Level": raid_card1_level,
                            "RAID Card1 Capacity": raid_card1_capacity,
                            "RAID Card2 Model": raid_card2_model,
                            "RAID Card2 Cache": raid_card2_cache,
                            "RAID Card2 FW": raid_card2_fw,
                            "RAID Card2 NVDATA": raid_card2_nvdata,
                            "RAID Card2 JBOD": raid_card2_jbod,
                            "OCP1 SN": ocp1_sn,
                            "OCP1 Model": ocp1_model,
                            "OCP1 MAC": ocp1_mac,
                            "OCP1 FW": ocp1_fw,
                            "OCP2 SN": ocp2_sn,
                            "OCP2 Model": ocp2_model,
                            "OCP2 MAC": ocp2_mac,
                            "OCP2 FW": ocp2_fw,
                            "PSU SN": psu_sn,
                            "PSU Capacity": psu_capacity,
                            "PSU Count": psu_count,
                            "GPU SN": gpu_sn,
                            "GPU Model": gpu_model,
                            "GPU FW": gpu_fw,
                            "GPU Count": gpu_count,
                            "BIOS Version": bios_version,
                            "BMC Version": bmc_version,
                            "MB CPLD Version": mb_cpld_version,
                            "BP1 CPLD Version": bp1_cpld_version,
                            "BP2 CPLD Version": bp2_cpld_version,
                            "VRD Version": vrd_version,
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

        # 압축 파일이 아닌 경우
        else:
            print("no archive file in base directory.")
            exit()

if __name__ == "__main__":
    main()

# SB, SC PciePortEnable9 값 구분
# AMD 서버같인 레이드 카드 없는 경우에 대한 분기
# 공백으로 출력되는 문제 에러메세지가 출력되로고 수정