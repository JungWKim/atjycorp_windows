from pathlib import Path
import tarfile as tf
import openpyxl as opx
import pandas as pd
import re

# 사용자가 변경해야할 값들
folder = Path(r"C:\Users\lionm\Downloads\LogCollectResult\20251107162128\device_logs")
output_file = Path(r"C:\Users\lionm\Downloads\output_one_click_log.xlsx")

def main():

    # 원클릭로그 폴더 내의 모든 .tar.gz 파일을 순회
    for archived_filename in folder.glob("*.tar.gz"):
        if archived_filename.is_file():
            with tf.open(archived_filename, "r:gz") as tar:

                # ipinfo_info >> bmc ip 정보 추출
                ipinfo_info_tarinfo = tar.getmember("dump_info/RTOSDump/networkinfo/ipinfo_info")
                if ipinfo_info_tarinfo.isfile():
                    with tar.extractfile(ipinfo_info_tarinfo) as ipinfo_info_file:
                        ipinfo_info_content = ipinfo_info_file.read().decode('utf-8')
                        bmc_ip_search_by_regex = re.search('^IP Address\s+:\s+([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)', ipinfo_info_content, re.MULTILINE)
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
                        print(server_config_txt_content)

                        # CPU 정보 추출
                        d

                        # PSU 정보 추출
                        psu_count = 0
                        psu_info_regex = re.search(r"([-]+PSU\s*info.*)(?=\n\s*\n-*Storage\s*info)", server_config_txt_content, re.DOTALL)
                        if psu_info_regex:
                            psu_info = psu_info_regex.group(1)
                            for line in psu_info.splitlines():

                                # PSU1 정보 추출
                                if line.startswith("1"):
                                    psu_count += 1
                                    psu1_info = line
                                    psu1_sn = psu1_info.split()[8]
                                    psu1_capacity = psu1_info.split()[13]

                                # PSU2 정보 추출
                                if line.strip().startswith("2"):
                                    psu_count += 1
                                    psu2_info = line
                                    psu2_sn = psu2_info.split()[8]
                                    psu2_capacity = psu2_info.split()[13]

                            # psu1 info 탐색하지 못한 경우
                            if not psu1_info:
                                psu1_sn = "psu1 info not found in server_config.txt"
                                psu1_capacity = "psu1 info not found in server_config.txt"

                            # psu2 info 탐색하지 못한 경우
                            if not psu2_info:
                                psu2_sn = "psu2 info not found in server_config.txt"
                                psu2_capacity = "psu2 info not found in server_config.txt"

                        # psu info 섹션 자체를 찾지 못한 경우
                        else:
                            psu1_sn = "psu info not found in server_config.txt"
                            psu1_capacity = "psu info not found in server_config.txt"
                            psu2_sn = "psu info not found in server_config.txt"
                            psu2_capacity = "psu info not found in server_config.txt"
                # server_config.txt 파일 자체를 찾지 못한 경우
                else:
                    psu1_sn = "no server_config.txt file"
                    psu1_capacity = "no server_config.txt file"
                    psu2_sn = "no server_config.txt file"
                    psu2_capacity = "no server_config.txt file"

                print(psu1_sn)
                print(psu1_capacity)
                print(psu2_sn)
                print(psu2_capacity)
                print(psu_count)
                break

                # netcard_info.txt >> nic 정보 추출
                netcard_info_tarinfo = tar.getmember("dump_info/LogDump/netcard/netcard_info.txt")
                if netcard_info_tarinfo.isfile():
                    with tar.extractfile(netcard_info_tarinfo) as netcard_info_file:
                        netcard_info_content = netcard_info_file.read().decode('utf-8')
                        nic1_model_match_list = re.findall(r'ProductName\s*:([^\n\r]+)', netcard_info_content, re.MULTILINE)
                        nic1_model = nic1_model_match_list[-1] if nic1_model_match_list else "not found in netcard_info.txt"
                        nic1_vendor_match_list = re.findall(r'Manufacture\s*:([^\n\r]+)', netcard_info_content, re.MULTILINE)
                        nic1_vendor = nic1_vendor_match_list[-1] if nic1_vendor_match_list else "not found in netcard_info.txt"
                        nic1_mac_match_list = re.findall(r'^\s*MacAddr:([^\n\r]+)', netcard_info_content, re.MULTILINE)
                        nic1_mac = nic1_mac_match_list[-2] if nic1_mac_match_list else "not found in netcard_info.txt"
                        nic1_slot_match_list = re.findall(r'SlotId\s*:([^\n\r]+)', netcard_info_content, re.MULTILINE)
                        nic1_slot = nic1_slot_match_list[-1] if nic1_slot_match_list else "not found in netcard_info.txt"
                # netcard_info.txt 파일 자체를 찾지 못한 경우
                else:
                    nic1_model = "no netcard_info.txt file"
                    nic1_vendor = "no netcard_info.txt file"
                    nic1_mac = "no netcard_info.txt file"
                    nic1_slot = "no netcard_info.txt file"

                # app_revision.txt >> bios, bmc, cpld, vrd, pn 정보 추출
                app_revision_txt_tarinfo = tar.getmember("dump_info/RTOSDump/versioninfo/app_revision.txt")
                if app_revision_txt_tarinfo.isfile():
                    with tar.extractfile(app_revision_txt_tarinfo) as app_revision_txt_file:
                        app_revision_txt_content = app_revision_txt_file.read().decode('utf-8')
                        print(app_revision_txt_content)
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
            
if __name__ == "__main__":
    main()

# 압축파일이 파일로 인식되지 않은 경우에 대한 모든 값들의 기본값 설정 필요
# 탐색할 변수명을 전부 리스트로?
# 바이오스 값들은 딕셔러리로?
# 티콘서버도 구분
# nic가 2개인 경우에 대한 대처