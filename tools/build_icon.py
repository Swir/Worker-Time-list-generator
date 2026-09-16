from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter
ROOT=Path(__file__).resolve().parents[1]; OUT=ROOT/"assets"/"worker-time-list.ico"
def draw(size=512):
    img=Image.new("RGBA",(size,size),(5,10,24,255)); glow=Image.new("RGBA",img.size,(0,0,0,0)); gd=ImageDraw.Draw(glow); gd.rounded_rectangle((42,42,size-42,size-42),radius=92,outline=(69,200,255,185),width=24); img.alpha_composite(glow.filter(ImageFilter.GaussianBlur(18))); d=ImageDraw.Draw(img); d.rounded_rectangle((42,42,size-42,size-42),radius=92,fill=(8,23,45,255),outline=(68,111,238,255),width=14); d.rounded_rectangle((112,96,400,414),radius=34,fill=(242,249,255,255),outline=(86,205,255,255),width=10); d.rectangle((144,150,368,174),fill=(31,64,104,255));
    for y in (220,276,332): d.rounded_rectangle((144,y,188,y+32),radius=8,fill=(74,222,170,255)); d.line((216,y+16,356,y+16),fill=(37,86,128,255),width=10)
    d.ellipse((44,310,194,460),fill=(8,23,45,255),outline=(86,205,255,255),width=12); d.line((119,340,119,388),fill=(242,249,255,255),width=12); d.line((119,388,154,410),fill=(242,249,255,255),width=12); return img
if __name__=="__main__": OUT.parent.mkdir(parents=True,exist_ok=True); draw().save(OUT,format="ICO",sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)]); print(OUT)
