import os
import re
import glob
import tarfile
import pandas as pd
from openpyxl import load_workbook

# -----------------------------
# 1) 사용자 지정 "원하는 사양" (혹은 input()으로 받아도 됨)
# -----------------------------
WANTED_CPU = "Xeon(R) Silver 4410Y"                   # 예: CPU 모델
#WANTED_CPU = "AMD EPYC 9354"                   # 예: CPU 모델
#WANTED_CPU = "AMD EPY C 9124"                   # 예: CPU 모델
WANTED_MEM_MANUFACTURERS = ["Samsung", "micron"]  # 복수 조건: 메모리 제조사 리스트
WANTED_MEM_SIZE = "32G"                       # 예: 메모리 용량 (예: "8G", "16G", "32G" 등)
WANTED_MEM_COUNT = 4                          # 예: 메모리 슬롯 개수
WANTED_NETCARD_MODELS = ["XC386", "BCM57414", "BCM957414N4140C", "BCM57416", "XC331", "XP330"]  # 복수 조건: 원하는 넷카드 모델 리스트

# device_logs 디렉토리 경로 (실제 경로로 수정)
base_dir = r"C:\Program Files\FusionServerTools\tmp\LogCollect_20251013013227\LogCollectResult\20251013133243\device_logs"

# 검수 결과 엑셀파일 저장 경로
output_excel = r"D:\log\hw_fw_inspection_result.xlsx"

# base_dir 하위의 모든 tar.gz 파일을 대상으로 함
tar_files = glob.glob(os.path.join(base_dir, "*.tar.gz"))

result_data = []

for tar_path in tar_files:
    basename = os.path.basename(tar_path)
    # 파일명 내에서 IP 주소 추출 
    ip_search = re.search(r"(\d{1,3}(?:\.\d{1,3}){3})", basename)
    if ip_search:
        ip = ip_search.group(1)
    else:
        continue  # IP 추출 실패시 건너뜀

    # 서버 시리얼 추출: 파일명 형식 예시
    # bmc_10.110.135.91_210619EYNVXGR2000008_20250314115620403.tar
    serial_search = re.search(r"bmc_\d{1,3}(?:\.\d{1,3}){3}_(.*?)_", basename)
    if serial_search:
        server_serial = serial_search.group(1)
    else:
        server_serial = ""
    
    # 초기값 설정
    cpld_ver = ""
    bios_ver = ""
    ibmc_ver = ""
    vrd_ver = ""
    bp1_cpld_ver = ""
    bp2_cpld_ver = ""
    bp3_cpld_ver = ""
    
    # 실제 추출된 CPU/Memory 값 (server_config.txt에서 추출)
    actual_cpu_model = ""
    actual_cpu_core = ""
    actual_mem_manufacturer_list = []
    actual_mem_model = []
    actual_mem_size_list = []
    actual_mem_count = 0
    # netcard 모델은 netcard_info.txt에서 추출할 예정.
    
    # 비교 후 최종 표기할 변수
    cpu_model = ""       
    cpu_core = ""       
    memory_manufacturer = ""
    memory_model = ""
    memory_size = ""
    memory_count = 0
    
    # netcard 관련 결과 (각 슬롯별)
    netcard1 = ""
    netcard1_version = ""
    slot1_mac = ""
    netcard2 = ""
    netcard2_version = ""
    slot2_mac = ""
    
    # NVME 정보 변수 초기화
    nvme_fw = ""
    nvme_capacity = ""
    
    # netcard_info.txt에서 추출한 슬롯 정보 저장할 딕셔너리 초기화
    slot1_info = {}
    slot2_info = {}
    
    try:
        with tarfile.open(tar_path, "r:gz") as tar:
            # --- app_revision.txt 파싱 (CPLD, iBMC, BIOS 버전) ---
            try:
                member = tar.getmember("dump_info/RTOSDump/versioninfo/app_revision.txt")
                with tar.extractfile(member) as f:
                    content = f.read().decode("utf-8", errors="ignore")
                    m = re.search(r"CPLD\s+Version:\s+(\S+)", content)
                    if m:
                        cpld_ver = m.group(1)
                    m = re.search(r"Active iBMC\s+Version:\s+(\S+)", content)
                    if m:
                        ibmc_ver = m.group(1)
                    m = re.search(r"Active BIOS\s+Version:\s+(\S+)", content)
                    if m:
                        bios_ver = m.group(1)
                    m = re.search(r"Mainboard\s+VRD:\s+(\S+)", content)
                    if m:
                        vrd_ver = m.group(1)
                    m = re.search(r"Disk\s+BP1\s+CPLD\s+Version:\s+(\S+)", content)
                    if m:
                        bp1_cpld_ver = m.group(1)
                    m = re.search(r"Disk\s+BP2\s+CPLD\s+Version:\s+(\S+)", content)
                    if m:
                        bp2_cpld_ver = m.group(1)
                    m = re.search(r"Disk\s+BP3\s+CPLD\s+Version:\s+(\S+)", content)
                    if m:
                        bp3_cpld_ver = m.group(1)
            except KeyError:
                pass

            # --- server_config.txt 파싱 (CPU, Memory, NVME 정보) ---
            try:
                member = tar.getmember("dump_info/RTOSDump/versioninfo/server_config.txt")
                with tar.extractfile(member) as f:
                    lines = f.read().decode("utf-8", errors="ignore").splitlines()
                    
                    # (1) CPU 정보 추출
                    cpu_info = []
                    cpu_core = []
                    for line in lines:
                        if line.strip().lower().startswith("cpu"):
                            parts = [p.strip() for p in line.split(",")]
                            if len(parts) >= 3:
                                cpu_info.append(parts[2])
                                cpu_core.append(parts[4])
                    if cpu_info:
                        actual_cpu_model = " / ".join(cpu_info)
                        actual_cpu_core = " / ".join(cpu_core)
                    
                    # (2) Memory 정보 추출
                    for line in lines:
                        if line.strip().startswith("Memory"):
                            parts = [p.strip() for p in line.split(",")]
                            if len(parts) >= 5:
                                mem_manu = parts[3]
                                if mem_manu.lower() == "unknown":
                                    continue
                                size_str = parts[4]
                                size_str = re.sub(r"\s*MB\s*", "", size_str, flags=re.IGNORECASE).strip()
                                try:
                                    size_mb = int(size_str)
                                    if size_mb % 1024 == 0:
                                        size_g = f"{size_mb // 1024}G"
                                    else:
                                        size_g = f"{round(size_mb/1024, 1)}G"
                                except:
                                    size_g = ""
                                mem_clock = parts[5].split(" ")[0]
                                mem_generation = parts[7]
                                mem_model = mem_generation + "-" + mem_clock
                                actual_mem_manufacturer_list.append(mem_manu)
                                actual_mem_model = mem_model
                                actual_mem_size_list.append(size_g)
                                actual_mem_count += 1
                    
                    # (3) NVME 정보 추출  
                    storage_text = "\n".join(lines)
                    nvme_match = re.search(r"ID\s*:\s*6.*?Firmware Version\s*:\s*(\S+).*?Capacity\s*:\s*([\d\.]+\s*\S+)", storage_text, re.DOTALL)
                    if nvme_match:
                        nvme_fw = nvme_match.group(1)
                        nvme_capacity = nvme_match.group(2)
            except KeyError:
                pass

            # --- netcard_info.txt 파싱 ---
            try:
                member = tar.getmember("dump_info/LogDump/netcard/netcard_info.txt")
                with tar.extractfile(member) as f:
                    netcard_content = f.read().decode("utf-8", errors="ignore")
                    # 블록 단위로 분리 (빈 줄 기준)
                    blocks = re.split(r'\n\s*\n', netcard_content)
                    netcard_entries = {}
                    for block in blocks:
                        product_match = re.search(r"ProductName\s*:\s*(\S+)", block)
                        firmware_match = re.search(r"FirmwareVersion\s*:\s*(\S+)", block)
                        slotid_match = re.search(r"SlotId\s*:\s*(\d+)", block)
                        port0_mac_match = re.search(r"Port0\s+.*?MacAddr:\s*([0-9A-Fa-f:]{17})", block, re.DOTALL)
                        if slotid_match:
                            slot = slotid_match.group(1)
                            netcard_entries[slot] = {
                                'ProductName': product_match.group(1) if product_match else "",
                                'FirmwareVersion': firmware_match.group(1) if firmware_match else "",
                                'Port0MAC': port0_mac_match.group(1) if port0_mac_match else ""
                            }
                    # 슬롯 정보가 없으면 기본 빈 딕셔너리 사용
                    slot1_info = netcard_entries.get('1', {})
                    slot2_info = netcard_entries.get('2', {})
            except KeyError:
                pass
            except Exception as e:
                print(f"파일 열기 실패 (netcard_info): {tar_path}\n에러: {e}")
                continue
    except Exception as e:
        print(f"tar 파일 열기 실패: {tar_path}\n에러: {e}")
        continue

    # -----------------------------
    # 2) 조건 비교 (원하는 값 vs 실제 추출)
    # -----------------------------
    # CPU 조건 비교
    if actual_cpu_model and (WANTED_CPU in actual_cpu_model):
        cpu_model = actual_cpu_model
        cpu_core = actual_cpu_core
    else:
        cpu_model = "false"
        cpu_core = "false"
    
    # Memory 조건 비교: 실제 메모리 제조사가 모두 WANTED_MEM_MANUFACTURERS 중 하나이고 용량이 동일해야 함
    if actual_mem_count == WANTED_MEM_COUNT:
        all_manu_match = all(m.strip().lower() in [x.lower() for x in WANTED_MEM_MANUFACTURERS] for m in actual_mem_manufacturer_list)
        all_size_match = all(s.strip().lower() == WANTED_MEM_SIZE.lower() for s in actual_mem_size_list)
        if all_manu_match and all_size_match:
            unique_manu = {m.strip().lower() for m in actual_mem_manufacturer_list}
            if len(unique_manu) == 1:
                mem_manufacturer = actual_mem_manufacturer_list[0].strip()
            else:
                mem_manufacturer = "/".join(unique_manu)
            memory_manufacturer = mem_manufacturer
            memory_model = actual_mem_model
            memory_size = WANTED_MEM_SIZE
            memory_count = WANTED_MEM_COUNT
        else:
            memory_manufacturer = "false"
            memory_size = "false"
            memory_count = "false"
    else:
        memory_manufacturer = "false"
        memory_size = "false"
        memory_count = "false"
    
    # Netcard 조건 비교 (슬롯별)
    if slot1_info and slot1_info.get('ProductName'):
        if slot1_info['ProductName'].upper() in [x.upper() for x in WANTED_NETCARD_MODELS]:
            netcard1 = slot1_info['ProductName']
            netcard1_version = slot1_info['FirmwareVersion']
            slot1_mac = slot1_info['Port0MAC'].lower()
        else:
            netcard1 = "false"
            netcard1_version = "false"
            slot1_mac = ""
    else:
        netcard1 = ""
        netcard1_version = ""
        slot1_mac = ""
    
    if slot2_info and slot2_info.get('ProductName'):
        if slot2_info['ProductName'].upper() in [x.upper() for x in WANTED_NETCARD_MODELS]:
            netcard2 = slot2_info['ProductName']
            netcard2_version = slot2_info['FirmwareVersion']
            slot2_mac = slot2_info['Port0MAC'].lower()
        else:
            netcard2 = "false"
            netcard2_version = "false"
            slot2_mac = ""
    else:
        netcard2 = ""
        netcard2_version = ""
        slot2_mac = ""
    
    result_data.append({
        "IP": ip,
        "서버 시리얼": server_serial,
        "CPU model": cpu_model,
        "CPU core": cpu_core,
        "Memory manufacturer": memory_manufacturer,
        "Memory model": memory_model,
        "Memory size": memory_size,
        "Memory count": memory_count,
        "NVME Firmware": nvme_fw,
        "NVME Capacity": nvme_capacity,
        "Netcard1": netcard1,
        "Slot1MAC": slot1_mac,
        "Netcard1버전": netcard1_version,
        "Netcard2": netcard2,
        "Netcard2버전": netcard2_version,
        "Slot2MAC": slot2_mac,
        "BIOS버전": bios_ver,
        "iBMC버전": ibmc_ver,
        "CPLD버전": cpld_ver,
        "VRD": vrd_ver,
        "BP1 CPLD버전": bp1_cpld_ver,
        "BP2 CPLD버전": bp2_cpld_ver,
        "BP3 CPLD버전": bp3_cpld_ver
    })

df = pd.DataFrame(result_data)

def ip_key(ip_str):
    return [int(part) for part in ip_str.split(".")]

if not df.empty and "IP" in df.columns:
    df = df.sort_values(by="IP", key=lambda col: col.map(ip_key))
else:
    print("DataFrame에 'IP' 컬럼이 없습니다.")

cols = [
    "IP",
    "서버 시리얼",
    "CPU model",
    "CPU core",
    "Memory manufacturer",
    "Memory model",
    "Memory size",
    "Memory count",
    "NVME Firmware",
    "NVME Capacity",
    "Netcard1",
    "Slot1MAC",
    "Netcard1버전",
    "Netcard2",
    "Netcard2버전",
    "Slot2MAC",
    "BIOS버전",
    "iBMC버전",
    "CPLD버전",
    "VRD",
    "BP1 CPLD버전",
    "BP2 CPLD버전",
    "BP3 CPLD버전"
]
df = df[cols]

df.to_excel(output_excel, index=False)
print("엑셀 파일 생성 완료:", output_excel)

# --- 오토핏: openpyxl을 이용하여 각 열 너비 조정 ---
wb = load_workbook(output_excel)
ws = wb.active

for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter
    for cell in col:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    adjusted_width = max_length + 2  # 약간의 여유를 둠
    ws.column_dimensions[col_letter].width = adjusted_width

wb.save(output_excel)
print("엑셀 파일 오토핏 적용 완료:", output_excel)
