import pandas as pd
from datetime import datetime
import os

class AttendanceSystem:
    def __init__(self):
        self.data_file = 'attendance_data.xlsx'
        self.members_file = 'members_list.txt'
        self.departments = ['락킹', '왁킹', '힙합', '걸스힙합', '하우스', '브레이킹']
        self.initialize_data_file()
        self.initialize_members_file()

    def initialize_data_file(self):
        if not os.path.exists(self.data_file):
            df = pd.DataFrame(columns=['날짜', '이름', '부서', '출석상태', '비고'])
            df.to_excel(self.data_file, index=False)

    def initialize_members_file(self):
        if not os.path.exists(self.members_file):
            with open(self.members_file, 'w', encoding='utf-8') as f:
                f.write("")

    def get_members_list(self):
        try:
            with open(self.members_file, 'r', encoding='utf-8') as f:
                members = {}
                for line in f:
                    if line.strip():
                        name, dept = line.strip().split(',')
                        members[name.strip()] = dept.strip()
            return members
        except FileNotFoundError:
            return {}

    def add_member(self, name, department):
        if department not in self.departments:
            return False, "존재하지 않는 부서입니다."
            
        members = self.get_members_list()
        if name not in members:
            with open(self.members_file, 'a', encoding='utf-8') as f:
                f.write(f"{name},{department}\n")
            return True, f"{name}님이 {department} 부서에 추가되었습니다."
        return False, f"{name}님은 이미 동아리원 목록에 있습니다."

    def remove_member(self, name):
        members = self.get_members_list()
        if name in members:
            del members[name]
            with open(self.members_file, 'w', encoding='utf-8') as f:
                for name, dept in members.items():
                    f.write(f"{name},{dept}\n")
            return True
        return False

    def check_attendance(self, names, status='출석'):
        today = datetime.now().strftime('%Y-%m-%d')
        df = pd.read_excel(self.data_file)
        
        # 이름 리스트 정리 (쉼표나 공백으로 구분)
        name_list = [name.strip() for name in names.replace(',', ' ').split() if name.strip()]
        
        if not name_list:
            return "입력된 이름이 없습니다."
        
        # 유효한 동아리원 목록 가져오기
        valid_members = self.get_members_list()
        
        results = []
        invalid_names = []
        
        # 먼저 모든 이름의 유효성 검사
        for name in name_list:
            if name not in valid_members:
                invalid_names.append(name)
        
        if invalid_names:
            return f"다음 이름은 동아리원 목록에 없습니다: {', '.join(invalid_names)}\n올바른 이름을 입력해주세요."
        
        # 모든 이름이 유효한 경우에만 출석 처리
        for name in name_list:
            # 오늘 날짜의 출석 기록이 있는지 확인
            if len(df[(df['날짜'] == today) & (df['이름'] == name)]) > 0:
                results.append(f"{name}님은 이미 오늘 출석 기록이 있습니다.")
                continue
            
            new_record = pd.DataFrame({
                '날짜': [today],
                '이름': [name],
                '부서': [valid_members[name]],
                '출석상태': [status],
                '비고': ['']
            })
            
            df = pd.concat([df, new_record], ignore_index=True)
            results.append(f"{name}님의 출석이 기록되었습니다. (상태: {status})")
        
        df.to_excel(self.data_file, index=False)
        return "\n".join(results)

    def get_attendance_summary(self, name=None, department=None):
        df = pd.read_excel(self.data_file)
        members = self.get_members_list()
        
        if name:
            if name not in members:
                return f"{name}님은 동아리원 목록에 없습니다."
            df = df[df['이름'] == name]
        elif department:
            if department not in self.departments:
                return f"존재하지 않는 부서입니다."
            df = df[df['부서'] == department]
        
        if len(df) == 0:
            return "출석 기록이 없습니다."
        
        if name:
            total_days = len(df)
            attendance_count = len(df[df['출석상태'] == '출석'])
            late_count = len(df[df['출석상태'] == '지각'])
            absent_count = len(df[df['출석상태'] == '결석'])
            
            attendance_rate = (attendance_count / total_days) * 100 if total_days > 0 else 0
            
            summary = f"\n=== {name}님의 출석 현황 ===\n"
            summary += f"부서: {members[name]}\n"
            summary += f"총 활동일수: {total_days}일\n"
            summary += f"출석: {attendance_count}일\n"
            summary += f"지각: {late_count}일\n"
            summary += f"결석: {absent_count}일\n"
            summary += f"출석률: {attendance_rate:.1f}%\n"
            
            return summary
        else:
            summary = f"\n=== {department} 부서 출석 현황 ===\n"
            for dept_member in [m for m, d in members.items() if d == department]:
                member_df = df[df['이름'] == dept_member]
                if len(member_df) > 0:
                    total_days = len(member_df)
                    attendance_count = len(member_df[member_df['출석상태'] == '출석'])
                    late_count = len(member_df[member_df['출석상태'] == '지각'])
                    absent_count = len(member_df[member_df['출석상태'] == '결석'])
                    attendance_rate = (attendance_count / total_days) * 100 if total_days > 0 else 0
                    
                    summary += f"\n{dept_member}님:\n"
                    summary += f"출석: {attendance_count}일\n"
                    summary += f"지각: {late_count}일\n"
                    summary += f"결석: {absent_count}일\n"
                    summary += f"출석률: {attendance_rate:.1f}%\n"
                else:
                    summary += f"\n{dept_member}님: 출석 기록 없음\n"
            
            return summary

    def view_attendance(self, date=None):
        df = pd.read_excel(self.data_file)
        
        if date:
            df = df[df['날짜'] == date]
            
        return df

    def save_attendance_to_csv(self, csv_file='attendance_data.csv'):
        """출결 데이터를 csv 파일로 저장"""
        df = pd.read_excel(self.data_file)
        df.to_csv(csv_file, index=False, encoding='utf-8-sig')
        return f"출결 데이터가 '{csv_file}' 파일로 저장되었습니다."

def main():
    system = AttendanceSystem()
    
    while True:
        print("\n=== 동아리 출결 관리 시스템 ===")
        print("1. 출석 체크")
        print("2. 출석 현황 조회")
        print("3. 날짜별 출석 조회")
        print("4. 동아리원 관리")
        print("5. 종료")
        print("6. 출결현황 csv로 저장")
        
        choice = input("\n원하는 작업을 선택하세요 (1-6): ")
        
        if choice == '1':
            print("\n이름을 입력하세요 (쉼표나 공백으로 구분):")
            print("예시: 홍길동, 김철수 이영희")
            names = input("> ")
            
            print("\n출석 상태를 선택하세요:")
            print("1. 출석")
            print("2. 지각")
            print("3. 결석")
            status_choice = input("선택 (1-3): ")
            
            status_map = {
                '1': '출석',
                '2': '지각',
                '3': '결석'
            }
            
            status = status_map.get(status_choice, '출석')
            print(system.check_attendance(names, status))
        
        elif choice == '2':
            print("\n조회 방식을 선택하세요:")
            print("1. 개인별 조회")
            print("2. 부서별 조회")
            view_choice = input("선택 (1-2): ")
            
            if view_choice == '1':
                name = input("조회할 이름을 입력하세요: ")
                print(system.get_attendance_summary(name=name))
            elif view_choice == '2':
                print("\n부서를 선택하세요:")
                for i, dept in enumerate(system.departments, 1):
                    print(f"{i}. {dept}")
                dept_choice = input("선택 (1-6): ")
                try:
                    dept = system.departments[int(dept_choice) - 1]
                    print(system.get_attendance_summary(department=dept))
                except (ValueError, IndexError):
                    print("잘못된 선택입니다.")
            else:
                print("잘못된 선택입니다.")
        
        elif choice == '3':
            date = input("조회할 날짜를 입력하세요 (YYYY-MM-DD): ")
            result = system.view_attendance(date)
            print("\n=== 출석 기록 ===")
            print(result)
        
        elif choice == '4':
            print("\n=== 동아리원 관리 ===")
            print("1. 동아리원 목록 보기")
            print("2. 동아리원 추가")
            print("3. 동아리원 삭제")
            sub_choice = input("선택 (1-3): ")
            
            if sub_choice == '1':
                members = system.get_members_list()
                if members:
                    print("\n=== 동아리원 목록 ===")
                    for dept in system.departments:
                        print(f"\n[{dept}]")
                        dept_members = [name for name, d in members.items() if d == dept]
                        if dept_members:
                            for member in dept_members:
                                print(f"- {member}")
                        else:
                            print("- 없음")
                else:
                    print("등록된 동아리원이 없습니다.")
            
            elif sub_choice == '2':
                name = input("추가할 동아리원 이름: ")
                print("\n부서를 선택하세요:")
                for i, dept in enumerate(system.departments, 1):
                    print(f"{i}. {dept}")
                dept_choice = input("선택 (1-6): ")
                try:
                    dept = system.departments[int(dept_choice) - 1]
                    success, message = system.add_member(name, dept)
                    print(message)
                except (ValueError, IndexError):
                    print("잘못된 선택입니다.")
            
            elif sub_choice == '3':
                name = input("삭제할 동아리원 이름: ")
                if system.remove_member(name):
                    print(f"{name}님이 동아리원 목록에서 삭제되었습니다.")
                else:
                    print(f"{name}님은 동아리원 목록에 없습니다.")
            
            else:
                print("잘못된 선택입니다.")
        
        elif choice == '5':
            print("프로그램을 종료합니다.")
            break
        
        elif choice == '6':
            message = system.save_attendance_to_csv()
            print(message)
        
        else:
            print("잘못된 선택입니다. 다시 시도해주세요.")

if __name__ == "__main__":
    main() 