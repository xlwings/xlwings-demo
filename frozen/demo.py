import xlwings as xw


def main():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello frozen xlwings!"


if __name__ == "__main__":
    main()

