"""Điểm khởi chạy: python main.py"""
from gui import App


def main():
    try:  # làm chữ sắc nét hơn trên Windows có màn hình độ phân giải cao
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()


if __name__ == "__main__":
    main()
