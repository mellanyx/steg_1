import docx
import MTK2
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX

doc = docx.Document('variant11.docx')


def run_get_spacing(run):
    rPr = run._r.get_or_add_rPr()
    spacings = rPr.xpath("./w:spacing")
    return spacings


def run_get_scale(run):
    rPr = run._r.get_or_add_rPr()
    scale = rPr.xpath("./w:w")
    return scale


def main():
    kod = ''
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            font_color = run.font.color.rgb
            font_size = run.font.size
            font_highlight_color = run.font.highlight_color
            font_scale = run_get_scale(run)
            font_spacing = run_get_spacing(run)

            if (font_color != RGBColor(0, 0, 0)):
                print('font_color')
                for i in range(len(run.text)):
                    kod += '1'
            elif (font_size.pt != 12.0):
                print('font_size')
                for i in range(len(run.text)):
                    kod += '1'
            elif (font_highlight_color != WD_COLOR_INDEX.WHITE):
                print('font_highlight_color')
                for i in range(len(run.text)):
                    kod += '1'
            elif (font_spacing):
                print('font_spacing')
                for i in range(len(run.text)):
                    kod += '1'
            elif (font_scale):
                print('font_scale')
                for i in range(len(run.text)):
                    kod += '1'
            else:
                for i in range(len(run.text)):
                    kod += '0'

    kod += "0000"
    print(kod)
    # print(len(kod))

    normaltext = MTK2.MTK2_decode(kod)
    print(normaltext)
    normaltext = bytes.fromhex(hex(int(kod, 2))[2:]).decode(encoding="koi8_r")
    print(normaltext)
    normaltext = bytes.fromhex(hex(int(kod, 2))[2:]).decode(encoding="cp866")
    print(normaltext)
    normaltext = bytes.fromhex(hex(int(kod, 2))[2:]).decode(encoding="cp1251")
    print(normaltext)

if __name__ == '__main__':
    main()