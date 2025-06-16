import pandas as pd

def get_user_inputs():
    """사용자로부터 필요한 입력값들을 받아옵니다."""
    excel_file = input("Excel 파일 경로를 입력하세요(예: 전산장비관리.xlsx): ")
    column_name = input("컬럼명을 입력하세요(예: IP): ")
    prefix_ip = input("IP 대역을 입력하세요(예: 211.218.228.): ")

    return excel_file, column_name, prefix_ip


def normalize_ip_prefix(prefix_ip):
    """IP 접두사를 정규화합니다. 마지막에 점이 없으면 추가합니다."""
    if not prefix_ip.endswith('.'):
        prefix_ip += '.'
    return prefix_ip


def extract_subnet_number(prefix_ip):
    """IP 접두사에서 서브넷 번호를 추출합니다."""
    return prefix_ip.rstrip('.').split('.')[-1]


def load_ip_data(excel_file, column_name):
    """Excel 파일에서 IP 데이터를 로드합니다."""
    try:
        df = pd.read_excel(excel_file)
        return df[column_name]
    except FileNotFoundError:
        print(f"파일을 찾을 수 없습니다: {excel_file}")
        return None
    except KeyError:
        print(f"컬럼을 찾을 수 없습니다: {column_name}")
        return None


def get_used_ips(ip_column, prefix_ip):
    """지정된 접두사로 시작하는 사용 중인 IP들을 반환합니다."""
    filtered_ips = ip_column[ip_column.str.startswith(prefix_ip, na=False)]
    return set(filtered_ips.dropna().astype(str))


def generate_all_possible_ips(prefix_ip):
    """주어진 접두사로 가능한 모든 IP 주소를 생성합니다."""
    return [f'{prefix_ip}{i}' for i in range(1, 255)]


def find_unused_ips(used_ips, all_possible_ips):
    """사용하지 않는 IP 목록을 찾습니다."""
    return [ip for ip in all_possible_ips if ip not in used_ips]


def save_results_to_excel(unused_ips, subnet_number):
    """결과를 Excel 파일로 저장합니다."""
    result_df = pd.DataFrame({'사용하지_않는_IP': unused_ips})
    output_file = f'사용하지않는IP목록_{subnet_number}대역.xlsx'

    try:
        result_df.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        print(f"파일 저장 중 오류가 발생했습니다: {e}")
        return None


def print_results(output_file, unused_count):
    """결과를 콘솔에 출력합니다."""
    if output_file:
        print(f"결과가 '{output_file}' 파일에 저장되었습니다.")
    print(f"총 사용하지 않는 IP 개수: {unused_count}개")


def main():
    """메인 실행 함수"""
    # 사용자 입력 받기
    excel_file, column_name, prefix_ip = get_user_inputs()

    # IP 접두사 정규화
    prefix_ip = normalize_ip_prefix(prefix_ip)

    # 데이터 로드
    ip_column = load_ip_data(excel_file, column_name)
    if ip_column is None:
        return

    # IP 분석
    used_ips = get_used_ips(ip_column, prefix_ip)
    all_possible_ips = generate_all_possible_ips(prefix_ip)
    unused_ips = find_unused_ips(used_ips, all_possible_ips)

    # 결과 저장 및 출력
    subnet_number = extract_subnet_number(prefix_ip)
    output_file = save_results_to_excel(unused_ips, subnet_number)
    print_results(output_file, len(unused_ips))


if __name__ == "__main__":
    main()
