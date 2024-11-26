import win32print
import win32ui
from PIL import Image, ImageWin, ImageDraw, ImageFont
import socket
import configparser
import os
import sys
import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ファイルの変更を監視するハンドラー
class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, server):
        self.server = server
        
    def on_modified(self, event):
        if event.src_path.endswith('.py'):
            print(f"ファイルが変更されました: {event.src_path}")
            self.server.restart()

# プリンターサーバーのメインクラス
class PrinterServer:
    def __init__(self):
        self.config = self._load_config()
        self.host = self.config.get('Server', 'host', fallback='localhost')
        self.port = self.config.getint('Server', 'port', fallback=5000)
        self.printer = Printer()
        self.running = True
        self.sock = None
        
    def restart(self):
        print("サーバーを再起動中...")
        self.running = False
        if self.sock:
            self.sock.close()
        python = sys.executable
        os.execl(python, python, *sys.argv)
    
    def start(self):
        # ファイル監視の設定
        event_handler = FileChangeHandler(self)
        observer = Observer()
        observer.schedule(event_handler, path='.', recursive=False)
        observer.start()
        
        try:
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            self.sock.bind((self.host, self.port))
            print(f"印刷サーバーが起動しました。待機中... ({self.host}:{self.port})")
            
            while self.running:
                try:
                    data, addr = self.sock.recvfrom(1024)
                    received_data = data.decode('utf-8').strip().split('|')
                    image_path = received_data[0]
                    text = received_data[1] if len(received_data) > 1 else ""
                    x_mm = float(received_data[2]) if len(received_data) > 2 else 60
                    y_mm = float(received_data[3]) if len(received_data) > 3 else 60
                    
                    print(f"受信: {addr}から")
                    print(f"画像パス: {image_path}")
                    print(f"テキスト: {text}")
                    print(f"x_mm: {x_mm}, y_mm: {y_mm}")
                    
                    self.printer.print_image_with_text(image_path, x_mm, y_mm, text)
                        
                except Exception as e:
                    print(f"エラーが発生しました: {e}")
                    
        except KeyboardInterrupt:
            print("サーバーを停止中...")
        finally:
            observer.stop()
            observer.join()
            if self.sock:
                self.sock.close()

    def _load_config(self):
        config = configparser.ConfigParser()
        config_path = 'Config.ini'
        
        if os.path.exists(config_path):
            config.read(config_path)
        else:
            config['Server'] = {
                'host': 'localhost',
                'port': '5000'
            }
            with open(config_path, 'w') as configfile:
                config.write(configfile)
        return config

# プリンター操作を行うクラス
class Printer:
    # プリンターの設定値
    WIDTH_IN_PIXEL = 8
    HEIGHT_IN_PIXEL = 10
    WIDTH_IN_MM = 4
    HEIGHT_IN_MM = 6
    LOGPIXELSX = 88
    LOGPIXELSY = 90
    PHYSICALWIDTH = 110
    PHYSICALHEIGHT = 111
    PHYSICALOFFSETX = 112
    PHYSICALOFFSETY = 113
    
    def trans_mm_to_pixel(self, hDC, x_mm, y_mm):
        # mmをピクセルに変換
        printable_area_px = hDC.GetDeviceCaps(self.WIDTH_IN_PIXEL), hDC.GetDeviceCaps(self.HEIGHT_IN_PIXEL)
        printable_area_mm = hDC.GetDeviceCaps(self.WIDTH_IN_MM), hDC.GetDeviceCaps(self.HEIGHT_IN_MM)
        px_px, py_px = printable_area_px
        px_mm, py_mm = printable_area_mm
        return (x_mm * px_px / px_mm, y_mm * py_px / py_mm)
        
    def add_text_to_image(self, image, text=""):
        # 画像にテキストを追加
        final_image = image.convert('RGBA')
        
        # フィルター画像の適用
        filter_path = './Assets/filter.png'
        if os.path.exists(filter_path):
            try:
                filter_img = Image.open(filter_path).convert('RGBA')
                filter_img = filter_img.resize(final_image.size)
                final_image = Image.alpha_composite(final_image, filter_img)
            except Exception as e:
                print(f"フィルター画像の処理でエラーが発生しました: {e}")
        
        final_image = final_image.convert('RGB')
        
        # テキストの追加処理
        if text.strip():
            draw = ImageDraw.Draw(final_image)
            try:
                font = ImageFont.truetype('./Assets/Font/Corporate-Logo-Bold-ver3.otf', 42)
            except Exception as e:
                print(f"フォントの読み込みに失敗しました: {e}")
                font = ImageFont.truetype('msgothic.ttc', 70)
            
            # 固定の文字間隔を設定
            char_spacing = 40  # 文字間隔を固定値で設定
            text_length = len(text)
            total_width = char_spacing * (text_length - 1)  # 全体の幅を計算
            
            # 開始位置を計算（中央揃え）
            start_x = (image.width - total_width) // 2
            baseline = image.height - 140
            
            # 文字を等間隔で配置
            for i, char in enumerate(text):
                x = start_x + (i * char_spacing)
                # anchorをleftに変更して左揃えで配置
                draw.text((x, baseline), char, font=font, fill=(0, 0, 0), anchor="ms")
        
        return final_image

    def print_image_with_text(self, image_path, x_mm, y_mm, text=""):
        # 画像の印刷処理
        printer_name = win32print.GetDefaultPrinter()
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)
        
        printer_margins = hDC.GetDeviceCaps(self.PHYSICALOFFSETX), hDC.GetDeviceCaps(self.PHYSICALOFFSETY)
        
        hDC.StartDoc(image_path)
        hDC.StartPage()
        
        # 画像の読み込みとテキスト追加
        bmp = Image.open(image_path)
        bmp = self.add_text_to_image(bmp, text)
            
        # 実行ファイルと同じディレクトリに保存するように変更
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'printed_images')
        os.makedirs(output_dir, exist_ok=True)
        
        base_name = os.path.splitext(os.path.basename(image_path))[0]
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(output_dir, f'{base_name}_{timestamp}.png')
        bmp.save(output_path)
        print(f"印刷用画像を保存しました: {output_path}")
            
        dib = ImageWin.Dib(bmp)
        scaled_width, scaled_height = self.trans_mm_to_pixel(hDC, x_mm, y_mm)

        print(f"元の画像サイズ: {bmp.size[0]}px x {bmp.size[1]}px")
        print(f"要求サイズ: {x_mm}mm x {y_mm}mm")
        print(f"変換後のピクセルサイズ: {scaled_width}px x {scaled_height}px")
        
        x1 = printer_margins[0]
        y1 = printer_margins[1]
        x2 = int(x1 + scaled_width)
        y2 = int(y1 + scaled_height)
        
        print(f"印刷位置: ({x1}, {y1}) -> ({x2}, {y2})")
        
        dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))
        
        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()

if __name__ == "__main__":
    server = PrinterServer()
    server.start()
