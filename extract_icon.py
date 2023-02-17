from PIL import Image
import win32gui  # pip package pywin32; has to be imported this way
import win32ui  # pip package pywin32; has to be imported this way
import win32api  # pip package pywin32; has to be imported this way

def extract_icon(exe_path: str, hwnd: int) -> Image: 
    width = height = 32
    """get the first resource icon from exefilename and return it as PIL Image"""
    ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
    try:
        large, small = win32gui.ExtractIconEx(exe_path, 0)

    except Exception as e:
        raise (e)

    if large: 
        wdc = win32gui.GetDC(hwnd)
        hdc = win32ui.CreateDCFromHandle(wdc)
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_x)
        hdc_2 = hdc.CreateCompatibleDC()
        hdc_2.SelectObject(hbmp)
        hdc_2.DrawIcon((0, 0), large[0])
        bmpstr = hbmp.GetBitmapBits(True)
        img = Image.frombuffer('RGBA', (width, height), bmpstr, 'raw', 'BGRA', 0, 1)
        win32gui.DestroyIcon(small[0])
        win32gui.DestroyIcon(large[0])

        hdc.DeleteDC()
        hdc_2.DeleteDC()
        win32gui.ReleaseDC(hwnd, wdc)
        win32gui.DeleteObject(hbmp.GetHandle())
        
        return img
    
    else:
        return None
