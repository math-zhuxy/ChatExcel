from excel_operator import excel_operate
def progress(i: int, s: str) -> None:
    pass

def test_operator():
    while True:
        function_call = input("你想要调用哪个函数？")
        user_input = input("输入你的要求：")
        if user_input == '-1':
            break
        ai_output, _ = excel_operate(
            user_input=user_input,
            file_path="./test.xlsx",
            function_called=function_call,
            progress_callback=progress
        )
        print(ai_output)