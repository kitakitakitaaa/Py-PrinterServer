import win32print
import win32ui
from PIL import Image, ImageWin, ImageDraw, ImageFont
import socket
import configparser
import os
import datetime

class PrinterServer:
    def __init__(self):
        self.config = self._load_config()
        self.host = self.config.get('Server', 'host', fallback='localhost')
        self.port = self.config.getint('Server', 'port', fallback=5000)
        self.printer = Printer()
    
    def _load_config(self):
        config = configparser.ConfigParser()
        config_path = 'Config.ini'
        
        if os.path.exists(config_path):
            config.read(config_path)
        else:
            # デフォルト設定の作成
            config['Server'] = {
                'host': 'localhost',
                'port': '5000'
            }
            with open(config_path, 'w') as configfile:
                config.write(configfile)
        
        return config
    
    def start(self):
        # UDPソケットの設定
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self.sock.bind((self.host, self.port))
        print(f"印刷サーバーが起動しました。待機中... ({self.host}:{self.port})")
        
        while True:
            try:
                # データの受信と分割
                data, addr = self.sock.recvfrom(1024)
                received_data = data.decode('utf-8').strip().split('|')
                image_path = received_data[0]
                text = received_data[1] if len(received_data) > 1 else ""
                
                print(f"受信: {addr}から")
                print(f"画像パス: {image_path}")
                print(f"テキスト: {text}")
                
                # テキストを追加した画像を作成して印刷
                self.printer.print_image_with_text(image_path, 76, 76, text)
                    
            except Exception as e:
                print(f"エラーが発生しました: {e}")

class Printer:
    # 下記はプリンタの固有情報を取得するためのキー
    # http://chokuto.ifdef.jp/urawaza/api/GetDeviceCaps.html
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
    
    
    # 印刷したいサイズ（mm）から画像サイズ（px）を導き出す
    def trans_mm_to_pixel(self, hDC, x_mm, y_mm):
        printable_area_px = hDC.GetDeviceCaps (self.WIDTH_IN_PIXEL), hDC.GetDeviceCaps (self.HEIGHT_IN_PIXEL)
        printable_area_mm = hDC.GetDeviceCaps (self.WIDTH_IN_MM), hDC.GetDeviceCaps (self.HEIGHT_IN_MM)

        px_px, py_px = printable_area_px
        px_mm, py_mm = printable_area_mm
        
        result = (x_mm * px_px / px_mm, y_mm * py_px / py_mm)
        
        return result
        
    def add_text_to_image(self, image, text=""):
        """画像にテキストとフィルター画像を追加する"""
        # 元画像をRGBA形式に変換
        final_image = image.convert('RGBA')
        
        # フィルター画像のパスをチェック
        filter_path = './Assets/filter.png'
        if os.path.exists(filter_path):
            try:
                filter_img = Image.open(filter_path).convert('RGBA')
                # フィルター画像のサイズを元画像のサイズに合わせてリサイズ
                filter_img = filter_img.resize(final_image.size)
                
                # 画像を合成
                final_image = Image.alpha_composite(final_image, filter_img)
                
            except Exception as e:
                print(f"フィルター画像の処理でエラーが発生しました: {e}")
        else:
            print("フィルター画像が見つかりません")
        
        # RGBモードに変換
        final_image = final_image.convert('RGB')
        
        # テキストが存在する場合のみテキストを描画
        if text.strip():  # 空白文字のみの場合もスキップ
            draw = ImageDraw.Draw(final_image)
            try:
                font = ImageFont.truetype('msgothic.ttc', 82)
            except:
                font = ImageFont.load_default()
            
            text_bbox = draw.textbbox((0, 0), text, font=font)
            text_width = text_bbox[2] - text_bbox[0]
            
            x = (image.width - text_width) // 2
            y = 20
            
            draw.text((x, y), text, font=font, fill='black')
        
        return final_image

    def print_image_with_text(self, image_path, x_mm, y_mm, text=""):
        """テキストを追加した画像を印刷する"""
        printer_name = win32print.GetDefaultPrinter()
        
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)
        
        # プリンター情報の取得
        printable_area = hDC.GetDeviceCaps(self.WIDTH_IN_PIXEL), hDC.GetDeviceCaps(self.HEIGHT_IN_PIXEL)
        printer_margins = hDC.GetDeviceCaps(self.PHYSICALOFFSETX), hDC.GetDeviceCaps(self.PHYSICALOFFSETY)
        
        hDC.StartDoc(image_path)
        hDC.StartPage()
        
        # 画像を開いてフィルター処理とテキストを追加
        bmp = Image.open(image_path)
        bmp = self.add_text_to_image(bmp, text)  # テキストの有無に関係なく処理を実行
            
        # 印刷用の画像を保存
        # output_dir = os.path.join(os.path.dirname(image_path), 'printed_images')
        # os.makedirs(output_dir, exist_ok=True)
        
        # base_name = os.path.splitext(os.path.basename(image_path))[0]
        # timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        # output_path = os.path.join(output_dir, f'{base_name}_{timestamp}.png')
        # bmp.save(output_path)
        # print(f"印刷用画像を保存しました: {output_path}")
            
        dib = ImageWin.Dib(bmp)
        scaled_width, scaled_height = self.trans_mm_to_pixel(hDC, x_mm, y_mm)

        # デバッグ用の出力
        print(f"元の画像サイズ: {bmp.size[0]}px x {bmp.size[1]}px")
        print(f"要求サイズ: {x_mm}mm x {y_mm}mm")
        print(f"変換後のピクセルサイズ: {scaled_width}px x {scaled_height}px")
        
        # 余白なしで印刷するための位置設定
        x1 = printer_margins[0]  # 左端
        y1 = printer_margins[1]  # 上端
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
