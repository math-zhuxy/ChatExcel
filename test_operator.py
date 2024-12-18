from excel_operator import excel_operate
def progress(i: int) -> None:
    pass

def test_operator():
    while True:
        user_input = input("输入你的要求：")
        if user_input == '-1':
            break
        ai_output = excel_operate(
            user_input=user_input,
            file_path="./test.xlsx",
            progress_callback=progress
        )
        print(ai_output)