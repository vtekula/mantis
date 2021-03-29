'''
Created Date: Tuesday, February 9th 2021
Author: vtekula

Copyright (c) 2021 vtekula
'''


import os
import re
import sys
from pathlib import Path
from io import BytesIO

import pandas as pd
from gooey import Gooey, GooeyParser
from PIL import Image, ImageDraw, ImageFont, ImageOps


OUTPUT_DIR = os.path.join(os.getcwd(), 'output')

if not os.path.isdir(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)


def add_border(args, iname, text):

    os.chdir(Path(iname).parent)
    iimg = Image.open(Path(iname).name)

    oimg = ImageOps.expand(iimg, border=args.border, fill=args.fill)
    W, H = oimg.size

    draw = ImageDraw.Draw(oimg)

    # imgfont = ImageFont.truetype('C:\Windows\Fonts\arial.ttf', args.fsize)

    file = open(r"C:\Windows\Fonts\arial.ttf", "rb")
    bytes_font = BytesIO(file.read())
    imgfont = ImageFont.truetype(bytes_font, args.fsize)

    w, h = draw.textsize(text, imgfont)
    offset_x, offset_y = imgfont.getoffset(text)
    w, h = w + offset_x, h + offset_y

    if args.top:
        x, y = (W-w)/2, (args.border-args.fsize)/2
    elif args.bottom:
        x, y = (W-w)/2, (H-h+(args.border-args.fsize)/2)

    draw.text((x, y), text, fill=args.fcolour, font=imgfont)

    oname = Path(iname).name.split('.')
    oname = oname[-2]+' output.'+oname[-1]

    os.chdir(OUTPUT_DIR)
    oimg.save(oname)


def extract_images_from_file(args):

    data = pd.read_excel(args.input)

    # change this if the name in the sheet are different
    cols = ['PATH', 'TEXT']
    df = pd.DataFrame(data, columns=cols)

    total = len(df)

    for i in range(len(df)):
        input_image, text = Path(df.iloc[i, 0]), df.iloc[i, 1]
        add_border(args, input_image, text)
        print("progress: {}/{}".format(i+1, total))


@Gooey(
    program_name="mantis",
    program_description="edit images in bulk",
    dump_build_config=False,
    show_restart_button=False,
    progress_regex=r"^progress: (\d+)/(\d+)$",
    progress_expr="x[0] / x[1] * 100",
    disable_progress_bar_animation=True
)
def main():

    parser = GooeyParser()

    paths = parser.add_argument_group("paths")

    paths.add_argument(
        "input",
        help="excel input file path",
        widget="FileChooser",
        gooey_options={
            'validator': {
                'test': "user_input.split('.')[-1] == 'xlsx'",
                'message': 'selected file is not an xlsx file'
            }
        }
    )

    image = parser.add_argument_group("image customization")

    image.add_argument(
        "border",
        help="border of the image",
        default=100,
        type=int,
        gooey_options={
            'validator': {
                'test': "int(user_input) > 0",
                'message': 'border should be greater than 0'
            }
        }
    )

    image.add_argument(
        "fill",
        help="hex value of the border (type/pick)",
        default="#000000",
        widget="ColourChooser",
        gooey_options={
            'validator': {
                'test': "re.search(r'^#(?:[0-9a-fA-F]{3}){1,2}$', user_input)",
                'message': 'hex is incorrect, format should be #000 or #000000'
            }
        }
    )

    font = parser.add_argument_group("font customization")

    font.add_argument(
        "fsize",
        help="font size, make sure it is less than the border for good results",
        default=90,
        type=int,
        gooey_options={
            'validator': {
                'test': "int(user_input) > 0",
                'message': 'font size should be greater than 0'
            }
        }
    )

    font.add_argument(
        "fcolour",
        help="font colour, hex value (type/pick)",
        default="#FFFFFF",
        widget="ColourChooser",
        gooey_options={
            'validator': {
                'test': "re.search(r'^#(?:[0-9a-fA-F]{3}){1,2}$', user_input)",
                'message': 'hex is incorrect, format should be #000 or #000000'
            }
        }
    )

    position = parser.add_mutually_exclusive_group(
        gooey_options={
            'initial_selection': 0
        }
    )
    position.add_argument(
        '--top',
        metavar='top',
        action='store_true',
        help='place the text on the top border'
    )
    position.add_argument(
        '--bottom',
        metavar='bottom',
        action='store_true',
        help='place the text on the bottom border'
    )

    args = parser.parse_args()

    extract_images_from_file(args)


if __name__ == '__main__':
    main()
