#!/usr/bin/env python3
"""Build the daily study sheet (.docx) for the 14-day Python plan.

Usage:
    python3 make_day.py N [--out DIR]     # N = 1..14
    python3 make_day.py --all [--out DIR]  # build every day (for testing)

Requires python-docx  (pip install python-docx).
"""
import argparse
import datetime
import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ----------------------------------------------------------------- palette
TEAL = "0F4C5C"
TEAL_LIGHT = "E6F0F2"
TEAL_MID = "9CC5CE"
AMBER = "E36414"
CODE_BG = "F3F5F7"
TEXT = "1F2933"
MUTED = "6B7280"
GRID = "D1D5DB"
FONT = "Calibri"
MONO = "Consolas"

START_DATE = datetime.date(2026, 11, 18)  # Day 1
VIDEO_URL = "https://www.youtube.com/watch?v=ix9cRaBkVe0"
WEEKDAYS = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"]
PAGE_W_CM = 17.0  # A4 (21cm) minus 2cm margins

# ------------------------------------------------------------------ content
DAYS = {
    1: dict(
        slug="Variables_Input",
        title="Biến, kiểu dữ liệu & input",
        video=("0:00:00", "0:51:46"),
        topics=["Giới thiệu Python, cài đặt", "Variables và type casting", "User input", "Mad Libs game", "Arithmetic và math"],
        concepts=[
            ("Biến (variable) và kiểu dữ liệu (type)",
             "Biến là cái tên trỏ tới một giá trị. Bốn kiểu cơ bản: int (số nguyên), float (số thực), str (chuỗi), bool (True/False). Dùng type() để kiểm tra.",
             'name = "Minh"\nage = 19\nheight = 1.7\nprint(type(name), type(age), type(height))\n# <class \'str\'> <class \'int\'> <class \'float\'>'),
            ("input() luôn trả về str",
             "Dù bạn gõ số 19, input() vẫn cho ra chuỗi \"19\". Muốn tính toán phải ép kiểu (type casting) bằng int() hoặc float() trước.",
             'age = input("Tuổi: ")   # gõ 19 -> "19" (str)\nprint(age + age)        # "1919" (nối chuỗi!)\nage = int(age)          # ép kiểu\nprint(age + 1)          # 20'),
            ("Toán tử số học",
             "Ngoài + - * /, hãy nhớ // (chia lấy phần nguyên), % (chia lấy dư) và ** (lũy thừa). Phép / luôn cho float.",
             "print(7 / 2)    # 3.5\nprint(7 // 2)   # 3\nprint(7 % 2)    # 1\nprint(2 ** 3)   # 8"),
        ],
        practice=[
            "Cài Python (python.org) và VS Code; chạy lệnh  python --version  để kiểm tra.",
            "Tạo file mad_libs.py: hỏi 4 từ (tính từ, danh từ, động từ, tên) rồi in ra một câu chuyện.",
            "Đổi câu chuyện thành phiên bản của riêng bạn (ít nhất 5 dòng).",
            "Thử int(\"12\") và int(\"abc\") để xem Python báo lỗi gì.",
        ],
        check=[
            "Vì sao input() với 5 và 3 rồi cộng lại cho ra \"53\" chứ không phải 8?",
            "7 // 2 và 7 % 2 cho kết quả bao nhiêu?",
        ],
    ),
    2: dict(
        slug="Conditions_Strings",
        title="Rẽ nhánh if & xử lý chuỗi",
        video=("0:51:46", "1:39:08"),
        topics=["If statements", "Calculator, weight và temperature converter", "Logical operators", "Conditional expressions", "String methods"],
        concepts=[
            ("if / elif / else và thụt lề (indentation)",
             "Python dùng thụt lề (4 spaces) để xác định khối lệnh. Chương trình chỉ chạy nhánh đầu tiên có điều kiện đúng.",
             'temp = 30\nif temp >= 35:\n    print("Nóng")\nelif temp >= 25:\n    print("Ấm")\nelse:\n    print("Mát")'),
            ("= (gán) khác == (so sánh); and / or / not",
             "Một dấu = để gán giá trị, hai dấu == để so sánh. Kết hợp điều kiện bằng and, or, not.",
             "x = 5                      # gán\nprint(x == 5)              # True\nprint(x > 3 and x < 10)    # True\nprint(not x == 5)          # False"),
            ("String methods trả về chuỗi mới (immutable)",
             "Chuỗi không thể sửa tại chỗ. Các method như upper(), strip(), replace() trả về một chuỗi mới; biến cũ không đổi trừ khi bạn gán lại.",
             'name = "minh"\nprint(name.upper())    # MINH\nprint(name)            # minh (không đổi)\nname = name.capitalize()\nprint(name)            # Minh'),
        ],
        practice=[
            "Viết tip calculator: nhập tổng hóa đơn và % tip, in ra số tiền tip và tổng phải trả.",
            "Viết password checker: hợp lệ khi dài >= 8 ký tự và không phải toàn chữ cái (gợi ý: .isalpha()).",
            "Làm lại 2 chương trình converter (cân nặng, nhiệt độ) trong video mà không nhìn code.",
        ],
        check=[
            "name.upper() có làm thay đổi biến name không? Vì sao?",
            "Khác nhau giữa = và == là gì? Cho một ví dụ.",
        ],
    ),
    3: dict(
        slug="Loops_Formatting",
        title="Vòng lặp & định dạng chuỗi",
        video=("1:39:08", "2:23:03"),
        topics=["String indexing", "Format specifiers", "While loops và compound interest", "For loops và countdown timer", "Nested loops"],
        concepts=[
            ("Indexing và slicing",
             "Chỉ số bắt đầu từ 0; chỉ số âm đếm từ cuối. Cú pháp s[start:stop] lấy từ start tới trước stop.",
             's = "Python"\nprint(s[0])      # P\nprint(s[-1])     # n\nprint(s[1:4])    # yth\nprint(s[::-1])   # nohtyP'),
            ("f-string và format specifier",
             "Đặt f trước chuỗi và viết biểu thức trong {}. Sau dấu : là định dạng: .2f (2 chữ số thập phân), , (dấu phẩy ngàn), > (căn phải).",
             'price = 1234.5678\nprint(f"{price:.2f}")     # 1234.57\nprint(f"{price:,.1f}")    # 1,234.6\nprint(f"{\'Hi\':>6}")       #     Hi'),
            ("while và for; vòng lặp lồng nhau",
             "for dùng khi biết số lần lặp; while dùng khi lặp tới khi điều kiện sai. Ở nested loop, vòng trong chạy hết một lượt cho mỗi lần vòng ngoài lặp.",
             'for i in range(1, 4):\n    for j in range(1, 4):\n        print(i * j, end=" ")\n    print()\n# 1 2 3\n# 2 4 6\n# 3 6 9'),
        ],
        practice=[
            "In bảng cửu chương từ 1 đến 9 bằng nested loop, căn cột đều bằng format specifier.",
            "Viết compound interest calculator dùng while loop (nhập vốn, lãi suất, số năm).",
            "Viết countdown timer bằng for loop và time.sleep(1).",
        ],
        check=[
            "Khi nào nên dùng while thay vì for?",
            "Với s = \"Python\", s[-3:] cho kết quả gì?",
        ],
    ),
    4: dict(
        slug="Collections",
        title="List, tuple, set & 2D collections",
        video=("2:23:03", "3:03:27"),
        topics=["Lists, sets, tuples", "Shopping cart program", "2D collections", "Quiz game"],
        concepts=[
            ("list vs tuple vs set",
             "list: có thứ tự, sửa được. tuple: có thứ tự, không sửa được. set: không thứ tự, không trùng lặp.",
             'fruits = ["apple", "banana"]    # list\nfruits.append("kiwi")\npoint = (3, 4)                  # tuple\nids = {1, 2, 2, 3}              # set\nprint(ids)                      # {1, 2, 3}'),
            ("Các thao tác list thường dùng",
             "append, remove, sort, len, toán tử in và slicing. Chỉ số bắt đầu từ 0 giống string.",
             "nums = [3, 1, 2]\nnums.sort()           # [1, 2, 3]\nprint(nums[0], len(nums))   # 1 3\nprint(2 in nums)      # True\nprint(nums[1:])       # [2, 3]"),
            ("2D list (list lồng list)",
             "Truy cập phần tử bằng grid[hàng][cột]. Duyệt bằng hai vòng for lồng nhau.",
             "grid = [[1, 2, 3],\n        [4, 5, 6]]\nprint(grid[1][2])     # 6\nfor row in grid:\n    for x in row:\n        print(x, end=' ')"),
        ],
        practice=[
            "Shopping cart: lưu sản phẩm dạng list các tuple (tên, giá), in hóa đơn và tổng tiền.",
            "Mở rộng quiz game lên 10 câu, lưu câu hỏi trong list of lists [câu hỏi, đáp án] và in điểm cuối cùng.",
            "Dùng set để loại các phần tử trùng trong một list bất kỳ.",
        ],
        check=[
            "Vì sao tuple không có append()? Khi nào nên dùng tuple?",
            "Cho grid ở trên, làm sao in tất cả phần tử của hàng đầu tiên?",
        ],
    ),
    5: dict(
        slug="Dict_Random",
        title="Dictionary & số ngẫu nhiên",
        video=("3:03:27", "3:52:12"),
        topics=["Dictionaries và concession stand", "Random numbers", "Number guessing game", "Rock, paper, scissors", "Dice roller"],
        concepts=[
            ("Dictionary: cặp key → value",
             "Truy cập bằng key. d[key] báo KeyError nếu key không tồn tại, còn d.get(key, mặc định) thì không.",
             'student = {"name": "Minh", "age": 19}\nprint(student["age"])          # 19\nprint(student.get("gpa", 0))   # 0\nstudent["gpa"] = 3.5           # thêm key mới'),
            ("Duyệt dictionary",
             "Dùng .items() để lấy cả key và value; dùng toán tử in để kiểm tra key.",
             'for k, v in student.items():\n    print(k, "->", v)\nprint("name" in student)   # True'),
            ("Module random",
             "random.randint(a, b) lấy số nguyên trong [a, b] (gồm cả hai đầu); random.choice(list) lấy một phần tử ngẫu nhiên.",
             'import random\nprint(random.randint(1, 6))\nprint(random.choice(["rock", "paper", "scissors"]))'),
        ],
        practice=[
            "Word counter: đếm số lần xuất hiện của mỗi từ trong một câu bằng dict.",
            "Number guessing game: máy chọn số 1-100, báo \"cao hơn/thấp hơn\" và đếm số lượt đoán.",
            "Rock-paper-scissors: chơi nhiều ván, lưu điểm vào dict {\"you\": 0, \"cpu\": 0}.",
        ],
        check=[
            "Khác nhau giữa d[\"x\"] và d.get(\"x\") là gì?",
            "random.randint(1, 6) có thể trả về 6 không?",
        ],
    ),
    6: dict(
        slug="Functions",
        title="Hàm (functions)",
        video=("3:52:12", "4:45:56"),
        topics=["Functions", "Default và keyword arguments", "*args và **kwargs", "Iterables", "Membership operators"],
        concepts=[
            ("def, parameter vs argument, return vs print",
             "Parameter là tên khai báo trong def; argument là giá trị bạn truyền khi gọi. return trả giá trị về để dùng tiếp, còn print chỉ in ra màn hình.",
             "def add(a, b):\n    return a + b\n\nresult = add(2, 3)\nprint(result * 10)   # 50"),
            ("Default và keyword arguments",
             "Tham số có giá trị mặc định thì có thể bỏ qua khi gọi. Keyword argument cho phép truyền theo tên, không cần đúng thứ tự.",
             'def greet(name, greeting="Xin chào"):\n    return f"{greeting}, {name}!"\n\nprint(greet("Minh"))\nprint(greet(greeting="Hi", name="An"))'),
            ("*args và **kwargs",
             "*args gom các argument vị trí thành một tuple; **kwargs gom các keyword argument thành một dict.",
             "def show(*args, **kwargs):\n    print(args, kwargs)\n\nshow(1, 2, x=3)   # (1, 2) {'x': 3}"),
        ],
        practice=[
            "Viết lại tip calculator và password checker của Day 2 thành các hàm calc_tip(bill, percent=10) và is_valid_password(pw).",
            "Viết hàm total(*numbers) cộng tất cả số được truyền vào.",
            "Dùng toán tử in để viết hàm kiểm tra một ký tự có phải nguyên âm không.",
        ],
        check=[
            "Hàm không có return thì trả về giá trị gì?",
            "show(1, 2, x=3) đưa mỗi argument vào args hay kwargs?",
        ],
    ),
    7: dict(
        slug="Comprehensions_Modules",
        title="Comprehension, module & scope (ôn tập tuần 1)",
        video=("4:45:56", "5:23:34"),
        topics=["List comprehensions", "Match-case statements", "Modules", "Scope resolution", "if __name__ == '__main__'"],
        concepts=[
            ("List comprehension",
             "Cú pháp [biểu_thức for x in iterable if điều_kiện] tạo list gọn hơn vòng for + append.",
             "squares = [x ** 2 for x in range(1, 6)]    # [1, 4, 9, 16, 25]\nnums = [1, 2, 3, 4, 5, 6]\nevens = [x for x in nums if x % 2 == 0]   # [2, 4, 6]"),
            ("import và module",
             "Mỗi file .py là một module. Dùng import để tái sử dụng code từ module chuẩn hoặc file do bạn tự viết.",
             "import math\nfrom random import randint\nprint(math.sqrt(16))\n# tạo file helpers.py rồi:  import helpers"),
            ("Scope và __name__ == \"__main__\"",
             "Biến trong hàm là local, biến ngoài hàm là global. Khối if __name__ == \"__main__\" chỉ chạy khi file được chạy trực tiếp, không chạy khi bị import.",
             'x = 10            # global\ndef f():\n    x = 5         # local, không đổi x global\n    return x\n\nif __name__ == "__main__":\n    print(f(), x)   # 5 10'),
        ],
        practice=[
            "Ôn tập: viết lại 3 vòng lặp bất kỳ của tuần này bằng list comprehension.",
            "Mini project: to-do list trong terminal với menu 1) Thêm 2) Xóa 3) Xem 4) Thoát.",
            "Dùng match-case để xử lý lựa chọn menu thay cho chuỗi if/elif.",
        ],
        check=[
            "Khi nào list comprehension không còn dễ đọc hơn vòng for thường?",
            "Nếu file A import file B, khối if __name__ == \"__main__\" của B có chạy không?",
        ],
    ),
    8: dict(
        slug="Projects_Banking",
        title="Dự án: Banking & Slot machine",
        video=("5:23:34", "6:07:26"),
        topics=["Banking program", "Slot machine", "Encryption program"],
        concepts=[
            ("Menu loop với while True và break",
             "Chương trình tương tác thường chạy trong vòng lặp vô hạn và thoát bằng break khi người dùng chọn Thoát.",
             'while True:\n    choice = input("1) Xem  2) Nạp  3) Rút  4) Thoát: ")\n    if choice == "4":\n        break'),
            ("Mỗi chức năng là một hàm, trả về trạng thái mới",
             "Thay vì sửa biến global, hãy để hàm nhận balance và return balance mới. Code dễ test và dễ đọc hơn.",
             'def deposit(balance):\n    amount = float(input("Số tiền nạp: "))\n    if amount <= 0:\n        print("Số tiền phải lớn hơn 0")\n        return balance\n    return balance + amount'),
            ("Chia nhỏ bài toán",
             "Trước khi code, viết ra các bước bằng tiếng Việt, rồi mỗi bước thành một hàm. Test từng hàm riêng trước khi ghép lại.",
             "# Slot machine, các bước:\n# 1. spin_row()   -> chọn 3 biểu tượng ngẫu nhiên\n# 2. print_row()  -> in kết quả\n# 3. get_payout() -> tính tiền thưởng"),
        ],
        practice=[
            "Làm lại banking program mà không nhìn video.",
            "Thêm chức năng xem lịch sử giao dịch (lưu vào một list).",
            "Slot machine: tự đổi bộ biểu tượng và bảng tiền thưởng.",
        ],
        check=[
            "Vì sao nên return số dư mới thay vì dùng biến global?",
            "Điều gì xảy ra nếu quên break trong while True?",
        ],
    ),
    9: dict(
        slug="OOP_Basics",
        title="OOP cơ bản: class & object",
        video=("6:07:26", "6:53:06"),
        topics=["Hangman game", "OOP intro", "Class variables"],
        concepts=[
            ("class, object, __init__ và self",
             "Class là bản thiết kế, object là sản phẩm tạo từ bản thiết kế đó. __init__ chạy khi tạo object; self đại diện cho chính object đang dùng.",
             'class Student:\n    def __init__(self, name, gpa):\n        self.name = name\n        self.gpa = gpa\n\n    def show(self):\n        print(f"{self.name}: {self.gpa}")\n\ns1 = Student("Minh", 3.5)\ns1.show()'),
            ("Object = dữ liệu + hành vi",
             "Mỗi object có dữ liệu (attributes) riêng. Đổi attribute của object này không ảnh hưởng object khác.",
             's2 = Student("An", 3.2)\ns1.gpa = 3.8\nprint(s1.gpa, s2.gpa)   # 3.8 3.2'),
            ("Class variable vs instance variable",
             "Class variable dùng chung cho mọi object (viết ngoài __init__). Instance variable (self.x) riêng cho từng object.",
             "class Student:\n    count = 0              # class variable\n    def __init__(self, name):\n        self.name = name   # instance variable\n        Student.count += 1"),
        ],
        practice=[
            "Tạo class Student (name, gpa, show()); tạo 3 sinh viên, đưa vào list và in người có GPA cao nhất.",
            "Thêm class variable đếm số Student đã tạo.",
            "Làm lại hangman: lưu các chữ đã đoán trong một set.",
        ],
        check=[
            "self là gì? Có thể đặt tên khác không?",
            "Class variable khác instance variable ở điểm nào?",
        ],
    ),
    10: dict(
        slug="Inheritance_Polymorphism",
        title="Kế thừa & đa hình",
        video=("6:53:06", "7:46:16"),
        topics=["Inheritance và multiple inheritance", "super()", "Polymorphism và duck typing", "Static methods", "Class methods"],
        concepts=[
            ("Kế thừa (inheritance) và super()",
             "Class con kế thừa mọi thứ của class cha (quan hệ \"là một\"). super() gọi lại code của class cha, thường dùng trong __init__.",
             "class GradStudent(Student):\n    def __init__(self, name, gpa, thesis):\n        super().__init__(name, gpa)\n        self.thesis = thesis"),
            ("Ghi đè (override) và đa hình (polymorphism)",
             "Class con định nghĩa lại method của cha. Cùng một lời gọi nhưng mỗi object phản ứng khác nhau; Python không cần cùng class cha, chỉ cần có method đó (duck typing).",
             'class Dog:\n    def speak(self): return "Gâu"\nclass Cat:\n    def speak(self): return "Meo"\n\nfor pet in [Dog(), Cat()]:\n    print(pet.speak())'),
            ("@staticmethod vs @classmethod",
             "staticmethod không nhận self hay cls, chỉ là hàm đặt trong class. classmethod nhận cls nên hay dùng để tạo object theo cách khác.",
             "class Student:\n    @staticmethod\n    def is_valid_gpa(g):\n        return 0 <= g <= 4\n\n    @classmethod\n    def from_string(cls, text):\n        name, gpa = text.split(\",\")\n        return cls(name, float(gpa))"),
        ],
        practice=[
            "Thêm class GradStudent kế thừa Student và override show() để in thêm thesis.",
            "Tạo thêm class Teacher, rồi viết vòng lặp gọi show() cho list gồm cả hai loại.",
            "Dùng from_string để tạo Student từ chuỗi \"Minh,3.5\".",
        ],
        check=[
            "super().__init__(...) làm gì và vì sao cần nó?",
            "Duck typing nghĩa là gì?",
        ],
    ),
    11: dict(
        slug="Magic_Property_Exceptions",
        title="Magic methods, property, decorator & xử lý lỗi",
        video=("7:46:16", "8:27:47"),
        topics=["Magic methods", "@property", "Decorators", "Exception handling", "File detection"],
        concepts=[
            ("__str__ và __repr__",
             "__str__ quyết định cách object hiển thị khi print() (cho người đọc); __repr__ dùng để debug.",
             'class Student:\n    def __init__(self, name, gpa):\n        self.name, self.gpa = name, gpa\n    def __str__(self):\n        return f"{self.name} ({self.gpa})"\n\nprint(Student("Minh", 3.5))   # Minh (3.5)'),
            ("@property: kiểm soát việc đọc và ghi attribute",
             "Dùng property để người ngoài vẫn viết s.gpa như bình thường nhưng bạn có thể kiểm tra giá trị hợp lệ.",
             "@property\ndef gpa(self):\n    return self._gpa\n\n@gpa.setter\ndef gpa(self, value):\n    if not 0 <= value <= 4:\n        raise ValueError(\"GPA phải từ 0 đến 4\")\n    self._gpa = value"),
            ("Decorator = hàm bọc hàm khác",
             "Decorator nhận một hàm và trả về hàm mới thêm tính năng (đo thời gian, ghi log...). Dùng cú pháp @tên_decorator.",
             "import time\n\ndef timer(func):\n    def wrapper(*args, **kwargs):\n        start = time.time()\n        result = func(*args, **kwargs)\n        print(f\"{time.time() - start:.3f}s\")\n        return result\n    return wrapper"),
            ("try / except / finally",
             "Bắt đúng loại lỗi cụ thể thay vì except trống. finally luôn chạy dù có lỗi hay không.",
             'try:\n    n = int(input("Số: "))\nexcept ValueError:\n    print("Không phải số hợp lệ")\nfinally:\n    print("Xong")'),
        ],
        practice=[
            "Thêm __str__ cho Student và in danh sách sinh viên.",
            "Dùng @property để chặn GPA ngoài khoảng 0-4.",
            "Viết decorator timer rồi áp dụng cho một hàm tính tổng 1..1.000.000.",
            "Bao input số của chương trình banking bằng try/except.",
        ],
        check=[
            "Vì sao nên bắt ValueError thay vì except: trống?",
            "Decorator thực chất làm gì với hàm được bọc?",
        ],
    ),
    12: dict(
        slug="Files_JSON_API",
        title="File, JSON & gọi API",
        video=("8:27:47", "9:22:19"),
        topics=["Writing và reading files", "Dates và times", "Alarm clock", "Multithreading (xem lướt)", "Requesting API data"],
        concepts=[
            ("with open(...) as f",
             "with tự đóng file khi xong. Chế độ: \"r\" đọc, \"w\" ghi đè, \"a\" ghi nối thêm. Nên thêm encoding=\"utf-8\" để tiếng Việt không bị lỗi.",
             'with open("todo.txt", "w", encoding="utf-8") as f:\n    f.write("Học Python\\n")\n\nwith open("todo.txt", "r", encoding="utf-8") as f:\n    print(f.read())'),
            ("JSON: lưu và đọc dữ liệu có cấu trúc",
             "json.dump ghi dict/list ra file; json.load đọc ngược lại. ensure_ascii=False giữ nguyên chữ có dấu.",
             'import json\ndata = {"name": "Minh", "tasks": ["a", "b"]}\nwith open("data.json", "w", encoding="utf-8") as f:\n    json.dump(data, f, ensure_ascii=False, indent=2)\nwith open("data.json", encoding="utf-8") as f:\n    print(json.load(f))'),
            ("Gọi API bằng requests",
             "Cài bằng  pip install requests. Luôn kiểm tra status_code (200 là thành công) và đặt timeout.",
             'import requests\nr = requests.get("https://api.github.com/users/octocat", timeout=10)\nif r.status_code == 200:\n    print(r.json()["name"])'),
        ],
        practice=[
            "Lưu to-do list của Day 7 ra file todo.json và đọc lại khi mở chương trình.",
            "Gọi API https://api.github.com/users/octocat và in ra một trường bất kỳ của JSON.",
            "Dùng datetime.now() để in ngày giờ hiện tại theo định dạng bạn thích.",
        ],
        check=[
            "Vì sao nên dùng with open(...) thay vì open() rồi close()?",
            "status_code 404 nghĩa là gì?",
        ],
    ),
    13: dict(
        slug="Venv_Git_Pytest",
        title="venv, Git/GitHub & pytest",
        video=None,
        topics=["Virtual environment, pip, requirements.txt", "Git và GitHub cơ bản", "Type hints, dataclasses, generators", "Kiểm thử với pytest"],
        concepts=[
            ("Virtual environment (venv)",
             "venv tạo môi trường riêng cho từng project để các thư viện không xung đột. requirements.txt giúp người khác cài lại đúng thư viện.",
             "python -m venv .venv\n.venv\\Scripts\\activate        # Windows\npip install requests\npip freeze > requirements.txt"),
            ("Git: add → commit → push",
             "add chọn thay đổi, commit lưu một mốc lịch sử, push gửi lên GitHub. Tạo repo rỗng trên GitHub trước, rồi nối vào máy.",
             'git init\ngit add .\ngit commit -m "Day 1-12 programs"\ngit remote add origin <URL repo>\ngit push -u origin main'),
            ("Type hints và pytest",
             "Type hints ghi chú kiểu dữ liệu (Python không bắt buộc khi chạy). Test là một hàm dùng assert để kiểm tra kết quả; chạy bằng lệnh pytest.",
             "def add(a: int, b: int) -> int:\n    return a + b\n\ndef test_add():\n    assert add(2, 3) == 5"),
        ],
        practice=[
            "Tạo repo GitHub, đưa toàn bộ chương trình tuần 1-2 vào đó cùng một file README ngắn.",
            "Tạo venv cho repo, cài requests và pytest, xuất requirements.txt.",
            "Viết 3 test cho một hàm bất kỳ của bạn và chạy pytest cho tới khi xanh.",
            "Đọc nhanh tài liệu về dataclasses và generators; thêm type hints cho 3 hàm.",
        ],
        check=[
            "Vì sao không nên cài thư viện thẳng vào Python hệ thống cho mọi project?",
            "Type hints có làm Python báo lỗi khi truyền sai kiểu không?",
        ],
    ),
    14: dict(
        slug="Final_Project",
        title="Dự án cuối: CSV → API → JSON",
        video=None,
        topics=["Tự thiết kế và hoàn thành một project nhỏ", "Ôn tập toàn bộ 14 ngày"],
        concepts=[
            ("Chia project thành các hàm nhỏ",
             "Mỗi hàm làm đúng một việc: đọc CSV, gọi API cho một dòng, ghi JSON. Hàm main() chỉ ghép các bước lại.",
             "def read_rows(path): ...\ndef fetch(row): ...\ndef save(results, path): ...\n\ndef main():\n    rows = read_rows(\"input.csv\")\n    results = [fetch(r) for r in rows]\n    save(results, \"output.json\")"),
            ("Xử lý lỗi từng dòng",
             "Một dòng lỗi hoặc một request thất bại không được làm cả chương trình dừng. Bắt lỗi cụ thể, ghi nhận và đi tiếp.",
             "for row in rows:\n    try:\n        results.append(fetch(row))\n    except requests.RequestException as e:\n        print(f\"Bỏ qua {row}: {e}\")"),
            ("README tốt",
             "Người lạ phải chạy được project chỉ bằng README: dự án làm gì, cách tạo venv, cài requirements, lệnh chạy và ví dụ đầu vào/đầu ra.",
             "# Tên dự án\nMô tả 1-2 câu.\n## Cài đặt\npython -m venv .venv\npip install -r requirements.txt\n## Chạy\npython main.py input.csv"),
        ],
        practice=[
            "Tạo file input.csv (5-10 dòng) và chọn một API miễn phí để gọi cho mỗi dòng.",
            "Viết code với ít nhất một class, một hàm có try/except, và lưu kết quả ra output.json.",
            "Tạo venv + requirements.txt, viết README, rồi push lên GitHub.",
            "Chạy lại toàn bộ từ repo sạch (clone mới) để chắc README đúng.",
        ],
        check=[
            "Project có chạy được trên máy khác chỉ bằng README không?",
            "Nếu API trả lỗi giữa chừng, chương trình xử lý thế nào?",
        ],
    ),
}


# ------------------------------------------------------------ low-level XML
def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def cell_borders(cell, **edges):
    """edges: top/left/bottom/right = (size_eighths_pt, hex) or None for nil."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        if edge not in edges:
            continue
        el = borders.find(qn("w:" + edge))
        if el is None:
            el = OxmlElement("w:" + edge)
            borders.append(el)
        spec = edges[edge]
        if spec is None:
            el.set(qn("w:val"), "nil")
        else:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(spec[0]))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), spec[1])


def cell_margins(cell, top=80, bottom=80, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = tcPr.find(qn("w:tcMar"))
    if mar is None:
        mar = OxmlElement("w:tcMar")
        tcPr.append(mar)
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = mar.find(qn("w:" + name))
        if el is None:
            el = OxmlElement("w:" + name)
            mar.append(el)
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    trPr.append(el)


def table_fixed(table, widths_cm):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tblPr = table._tbl.tblPr  # autofit=False already writes <w:tblLayout w:type="fixed"/>
    total = int(sum(widths_cm) / 2.54 * 1440)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")
    for row in table.rows:
        for idx, w in enumerate(widths_cm):
            row.cells[idx].width = Cm(w)
    grid = table._tbl.tblGrid
    for idx, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths_cm[idx] / 2.54 * 1440)))


def clear_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def para_border_bottom(p, hex_color=GRID, size=6, space=2):
    pPr = p._p.get_or_add_pPr()
    pb = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(size))
    b.set(qn("w:space"), str(space))
    b.set(qn("w:color"), hex_color)
    pb.append(b)
    pPr.append(pb)


def spacing(p, before=0, after=0, line=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing = line


def run(p, text, size=10.5, bold=False, italic=False, color=TEXT, font=FONT):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = RGBColor.from_string(color)
    r.font.name = font
    rPr = r._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), font)
    return r


def add_hyperlink(p, url, text, size=10.5, color="0B63CE"):
    part = p.part
    r_id = part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True
    )
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), FONT)
    rPr.append(rFonts)
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    rPr.append(c)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(sz)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    link.append(r)
    p._p.append(link)


def add_page_field(p, size=8.5, color=MUTED):
    def fld(kind):
        r = p.add_run()
        r.font.size = Pt(size)
        r.font.color.rgb = RGBColor.from_string(color)
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), kind)
        r._r.append(el)

    fld("begin")
    r = p.add_run()
    r.font.size = Pt(size)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    r._r.append(instr)
    fld("separate")
    run(p, "1", size=size, color=color)
    fld("end")



# ------------------------------------------------- schema element ordering
_ORDER = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
            "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
            "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
            "sectPr", "pPrChange"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
             "textDirection", "tcFitText", "vAlign", "hideMark"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
              "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
              "tblLook", "tblCaption", "tblDescription"],
}


def _local(tag):
    return tag.split("}")[1] if "}" in tag else tag


def fix_order(root):
    for name, order in _ORDER.items():
        for el in root.iter(qn("w:" + name)):
            kids = list(el)
            kids.sort(key=lambda k: order.index(_local(k.tag)) if _local(k.tag) in order else len(order))
            for k in kids:
                el.remove(k)
            for k in kids:
                el.append(k)


# ----------------------------------------------------------------- helpers
def hms_to_sec(s):
    parts = [int(x) for x in s.split(":")]
    while len(parts) < 3:
        parts.insert(0, 0)
    return parts[0] * 3600 + parts[1] * 60 + parts[2]


def pretty_hms(s):
    return s if s.count(":") == 2 else s


def first_para(cell):
    return cell.paragraphs[0]


def new_para(cell, first=False):
    return cell.paragraphs[0] if first else cell.add_paragraph()


def section_heading(doc, label):
    p = doc.add_paragraph()
    spacing(p, before=14, after=5)
    p.paragraph_format.keep_with_next = True
    run(p, label.upper(), size=10.5, bold=True, color=TEAL)
    para_border_bottom(p, hex_color=TEAL, size=8, space=3)
    return p


def spacer(doc, pts=4):
    p = doc.add_paragraph()
    spacing(p, 0, 0)
    p.paragraph_format.line_spacing = Pt(pts)
    return p


# ------------------------------------------------------------- doc sections
def build_header(doc, n, d):
    date = START_DATE + datetime.timedelta(days=n - 1)
    date_txt = f"{WEEKDAYS[date.weekday()]}, {date.strftime('%d/%m/%Y')}"
    t = doc.add_table(rows=1, cols=1)
    clear_table_borders(t)
    table_fixed(t, [PAGE_W_CM])
    c = t.cell(0, 0)
    shade(c, TEAL)
    cell_margins(c, top=220, bottom=220, left=300, right=300)

    p = first_para(c)
    spacing(p, 0, 2)
    run(p, "PYTHON IN 14 DAYS", size=8.5, bold=True, color=TEAL_MID)
    p = c.add_paragraph()
    spacing(p, 0, 0)
    run(p, f"DAY {n} / 14", size=11, bold=True, color="FFB27A")
    p = c.add_paragraph()
    spacing(p, 2, 4)
    run(p, d["title"], size=22, bold=True, color="FFFFFF")
    p = c.add_paragraph()
    spacing(p, 0, 0)
    run(p, date_txt, size=10, color="D5E6EA")
    return t


def build_progress(doc, n):
    spacer(doc, 6)
    t = doc.add_table(rows=1, cols=14)
    clear_table_borders(t)
    w = PAGE_W_CM / 14
    table_fixed(t, [w] * 14)
    for i in range(14):
        c = t.cell(0, i)
        day = i + 1
        if day < n:
            fill, fg, bold = TEAL_MID, "FFFFFF", False
        elif day == n:
            fill, fg, bold = AMBER, "FFFFFF", True
        else:
            fill, fg, bold = "E5E7EB", MUTED, False
        shade(c, fill)
        cell_borders(c, left=(18, "FFFFFF"), right=(18, "FFFFFF"))
        cell_margins(c, top=40, bottom=40, left=20, right=20)
        p = first_para(c)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        spacing(p, 0, 0)
        run(p, str(day), size=8.5, bold=bold, color=fg)


def build_video(doc, d):
    section_heading(doc, "1 · Video hôm nay")
    t = doc.add_table(rows=1, cols=1)
    clear_table_borders(t)
    table_fixed(t, [PAGE_W_CM])
    c = t.cell(0, 0)
    shade(c, TEAL_LIGHT)
    cell_borders(c, left=(36, TEAL))
    cell_margins(c, top=120, bottom=120, left=220, right=200)
    no_split(t.rows[0])
    p = first_para(c)
    spacing(p, 0, 2)
    if d["video"]:
        start, end = d["video"]
        minutes = round((hms_to_sec(end) - hms_to_sec(start)) / 60)
        run(p, "Bro Code · Python Full Course for free", size=11, bold=True, color=TEAL)
        p = c.add_paragraph()
        spacing(p, 0, 3)
        run(p, f"Xem từ {start} đến {end}", size=11, bold=True)
        run(p, f"   (khoảng {minutes} phút)", size=10.5, color=MUTED)
        p = c.add_paragraph()
        spacing(p, 0, 0)
        add_hyperlink(p, f"{VIDEO_URL}&t={hms_to_sec(start)}s", f"Mở video tại phút {start}")
    else:
        run(p, "Hôm nay không có video", size=11, bold=True, color=TEAL)
        p = c.add_paragraph()
        spacing(p, 0, 0)
        run(p, "Tự thực hành theo checklist bên dưới. Khi cần, tra cứu docs.python.org hoặc hỏi Claude.", size=10.5)


def build_topics(doc, d):
    section_heading(doc, "2 · Nội dung (topics)")
    for topic in d["topics"]:
        p = doc.add_paragraph(style="List Bullet")
        spacing(p, 0, 1)
        run(p, topic, size=10.5)


def build_concepts(doc, d):
    section_heading(doc, "3 · Cần nắm chắc")
    inner_w = PAGE_W_CM - 1.0 - 0.8  # card content width minus cell side margins
    for i, (title, text, code) in enumerate(d["concepts"], 1):
        t = doc.add_table(rows=1, cols=2)
        clear_table_borders(t)
        table_fixed(t, [1.0, PAGE_W_CM - 1.0])
        no_split(t.rows[0])

        num = t.cell(0, 0)
        shade(num, AMBER)
        cell_margins(num, top=100, bottom=100, left=40, right=40)
        p = first_para(num)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        spacing(p, 4, 0)
        run(p, str(i), size=16, bold=True, color="FFFFFF")

        body = t.cell(0, 1)
        shade(body, "FFFFFF")
        cell_borders(body, top=(6, GRID), right=(6, GRID), bottom=(6, GRID))
        cell_margins(body, top=110, bottom=110, left=200, right=200)
        p = first_para(body)
        spacing(p, 0, 2)
        run(p, title, size=11.5, bold=True, color=TEAL)
        p = body.add_paragraph()
        spacing(p, 0, 6)
        run(p, text, size=10.5)

        code_tbl = body.add_table(rows=1, cols=1)
        clear_table_borders(code_tbl)
        table_fixed(code_tbl, [inner_w])
        cc = code_tbl.cell(0, 0)
        shade(cc, CODE_BG)
        cell_borders(cc, left=(24, TEAL_MID))
        cell_margins(cc, top=90, bottom=90, left=160, right=160)
        for j, line in enumerate(code.split("\n")):
            cp = new_para(cc, first=(j == 0))
            spacing(cp, 0, 0)
            cp.paragraph_format.line_spacing = 1.0
            run(cp, line if line else " ", size=9.5, color="0B3C49", font=MONO)
        tail = body.add_paragraph()  # Word needs a paragraph after a nested table
        spacing(tail, 0, 0)
        tail.paragraph_format.line_spacing = Pt(2)
        spacer(doc, 7)


def box_line(doc, symbol, text, indent_cm=0.8):
    p = doc.add_paragraph()
    spacing(p, 0, 3)
    p.paragraph_format.left_indent = Cm(indent_cm)
    p.paragraph_format.first_line_indent = Cm(-indent_cm)
    run(p, symbol + "\u2002", size=12, color=TEAL, font="Segoe UI Symbol")
    run(p, text, size=10.5)
    return p


def build_practice(doc, d):
    section_heading(doc, "4 · Bài tập thực hành")
    for item in d["practice"]:
        box_line(doc, "\u2610", item)


def build_check(doc, d):
    section_heading(doc, "5 · Tự kiểm tra (trả lời không nhìn bài)")
    for i, q in enumerate(d["check"], 1):
        p = doc.add_paragraph()
        spacing(p, 2, 2)
        p.paragraph_format.keep_with_next = True
        run(p, f"{i}.  ", size=10.5, bold=True, color=AMBER)
        run(p, q, size=10.5)
        for _ in range(2):
            line = doc.add_paragraph()
            spacing(line, 6, 0)
            para_border_bottom(line, hex_color=GRID, size=4, space=1)


def build_notes(doc):
    section_heading(doc, "Ghi chú của mình")
    for _ in range(3):
        line = doc.add_paragraph()
        spacing(line, 8, 0)
        para_border_bottom(line, hex_color=GRID, size=4, space=1)


def build_footer(section, n):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    spacing(p, 0, 0)
    run(p, f"Python in 14 days  ·  Day {n}/14  ·  Trang ", size=8.5, color=MUTED)
    add_page_field(p)


def make_doc(n, out_dir):
    d = DAYS[n]
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin, sec.bottom_margin = Cm(1.6), Cm(1.8)

    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(10.5)
    rpr = st.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(attr), FONT)
    st.paragraph_format.space_after = Pt(0)

    doc.core_properties.title = f"Python Day {n}: {d['title']}"
    doc.core_properties.author = "Claude"

    build_header(doc, n, d)
    build_progress(doc, n)
    build_video(doc, d)
    build_topics(doc, d)
    build_concepts(doc, d)
    build_practice(doc, d)
    build_check(doc, d)
    build_notes(doc)
    build_footer(sec, n)

    fix_order(doc.element)
    fix_order(sec.footer._element)
    z = doc.settings.element.find(qn("w:zoom"))
    if z is not None:
        z.set(qn("w:percent"), "100")

    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"Day_{n:02d}_{d['slug']}.docx")
    doc.save(path)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("day", nargs="?", type=int, help="day number 1-14")
    ap.add_argument("--out", default=".", help="output folder")
    ap.add_argument("--all", action="store_true")
    a = ap.parse_args()
    if a.all:
        for n in sorted(DAYS):
            print(make_doc(n, a.out))
        return
    if a.day not in DAYS:
        sys.exit("day must be 1..14")
    print(make_doc(a.day, a.out))


if __name__ == "__main__":
    main()
