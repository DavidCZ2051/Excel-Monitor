import xlwings as xw
import pyautogui
import PIL

WIDTH = 1920 # width of the screenshot in pixels
HEIGHT = 1080 # height of the screenshot in pixels
AREA_SIZE = 20 # size of each area in pixels
# Areas are used to split the screenshot into smaller areas to get the average color of each area
# They are squares with sides of length AREA_SIZE
# Examples:  
# if AREA_SIZE = 1, then the output will be a WIDTH x HEIGHT pixel image
# if AREA_SIZE = 2, then the output will be a WIDTH/2 x HEIGHT/2 pixel image
SHEET_WIDTH = WIDTH // AREA_SIZE # width of the excel sheet in cells
SHEET_HEIGHT = HEIGHT // AREA_SIZE # height of the excel sheet in cells


def screenshot() -> None:
    print('screenshotting...')
    im = pyautogui.screenshot(region=(0, 0, WIDTH, HEIGHT))
    im.save('screenshot.png')


def split_screenshot() -> list:
    print('splitting...')
    # splits the screenshot into AREA_SIZE x AREA_SIZE pixel areas and gets the average color of each area
    areas = []  # each element in this list is a tuple which contains the average color of the area (r, g, b)
    im = PIL.Image.open('screenshot.png')
    for i in range(SHEET_WIDTH):
        for j in range(SHEET_HEIGHT):
            area = im.crop((i * AREA_SIZE, j * AREA_SIZE, (i + 1)
                           * AREA_SIZE, (j + 1) * AREA_SIZE))
            r, g, b = 0, 0, 0
            for x in range(AREA_SIZE):
                for y in range(AREA_SIZE):
                    r += area.getpixel((x, y))[0]
                    g += area.getpixel((x, y))[1]
                    b += area.getpixel((x, y))[2]
            r //= AREA_SIZE ** 2
            g //= AREA_SIZE ** 2
            b //= AREA_SIZE ** 2
            areas.append((r, g, b))
    return areas


def display(wb, areas) -> None:
    # writes each area into a cell in the excel sheet.
    print('displaying...')
    sheet = wb.sheets[0]
    for i in range(SHEET_WIDTH):
        for j in range(SHEET_HEIGHT):
            sheet.range((j + 1, i + 1)).color = areas[i * SHEET_HEIGHT + j]


def main() -> None:
    wb = xw.Book('main.xlsx')

    screenshot()
    areas = split_screenshot()
    display(wb, areas)


if __name__ == '__main__':
    main()
