import libxlpy as libxl

# 1. 查看整个 libxl 模块里有什么
print(help(libxl))
# 你会看到 Book, Format, Font, Sheet 等类

# 2. 创建一个核心对象，比如 Book
book = libxl.Book()

book.setKey("helloworld", "windows-24292a0f0ecde30c67bc6f61a8l6p9e3")

# 3. 查看 Book 对象有哪些方法
print(help(book))
# 你会看到 'add_sheet', 'add_format', 'save', 'load', 'release' 等方法
# 注意：以双下划线开头和结尾的（如 __init__）是 Python 的特殊方法

# 4. 创建一个 Sheet 对象并查看它的方法
sheet = book.addSheet('MySheet')
print(help(sheet))
# 你会看到 'write_str', 'write_num', 'read_str', 'set_col', 'row_height' 等方法

format = book.addFormat()
font = book.addFont()

print(help(format))
print(help(font))

format.setFillPattern(libxl.FILLPATTERN_SOLID)
format.setPatterForegroundColor(libxl.COLOR_RED)

format.setAlignH(libxl.ALIGNH_CENTER)
format.setAlignV(libxl.ALIGNV_CENTER)
format.setBorder(libxl.BORDERSTYLE_THIN)
format.setBorderColor(libxl.COLOR_BLACK)

# 应用样式
sheet.writeStr(1, 1, "Styled Text", format)

# 自动调整列宽
sheet.setCol(1, 1, 20)


book.save("aaa.xlsx")