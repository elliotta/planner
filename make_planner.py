#!/usr/bin/env python3
#from moon.dialamoon import Moon
from datetime import datetime
from datetime import timedelta
import tempfile
import os
from sys import platform
import logging
import xml.dom.minidom
import subprocess
import json
from PyPDF2 import PdfFileReader, PdfFileMerger

WEEKDAY_COLOR = '000000'
WEEKEND_COLOR = '5eaf0c'

WEEKLY_TEMPLATE = None
DAILY_TEMPLATE  = None
ALTERNATE_TEMPLATES = {}
WEEKLY_TEMPLATE_FILE = '/Users/elliotta/Documents/Planner/weekly.svg'
DAILY_TEMPLATE_FILE  = '/Users/elliotta/Documents/Planner/daily.svg'

MERGER = None
MERGER_PAGE_COUNT = -1 # It's ridiculous that I have to do this myself


'''
def get_moon(date=datetime.now(), size=(200,200), filename='moon_test'):
    moon = Moon(size=size)
    moon.set_moon_phase(date.strftime('%Y-%m-%d'))
    moon.image
    moon.save_to_disk(filename)
'''

def make_alterate_dailies_dict(json_file):
    logging.info('Assembling alternate templates from %s' % json_file)
    global ALTERNATE_TEMPLATES
    alternate_files = {}
    try:
        with open(json_file) as f:
            alternates = json.load(f)
    except Exception as exc:
        logging.error('Cannot load alternates file: ' % exc)
        raise
    for date_string, filename in alternates.items():
        date = datetime.strptime(date_string, '%Y%m%d').date()
        if filename not in alternate_files:
            alternate_files[filename] = xml.dom.minidom.parse(filename)
        ALTERNATE_TEMPLATES[date] = alternate_files[filename]


def svg2pdf(infile, outfile):
    logging.debug('Converting %s to %s' % (infile, outfile))
    if platform == 'darwin':
        subprocess.run(['/Applications/Inkscape.app/Contents/MacOS/inkscape', '--export-filename='+outfile, infile])
    else:
        subprocess.run(['inkscape', '--export-filename='+outfile, infile])


def add_page_to_merger(page_to_add, bookmark_name=None, bookmark_parent=None):
    logging.debug('Adding %s to merger' % page_to_add)
    global MERGER, MERGER_PAGE_COUNT
    if not MERGER:
        MERGER = PdfFileMerger()
    with open(page_to_add, 'rb') as f:
        MERGER.append(PdfFileReader(f))
    MERGER_PAGE_COUNT += 1
    if bookmark_parent:
        return MERGER.add_outline_item(bookmark_name, parent=bookmark_parent, pagenum=MERGER_PAGE_COUNT)
    if bookmark_name:
        return MERGER.addBookmark(bookmark_name, pagenum=MERGER_PAGE_COUNT)


def replace_fill_color(style_string, color_string):
    fill_prefix = 'fill:#'
    style_list = style_string.split(';')
    found_fill = False
    for i, style in enumerate(style_list):
        if style.startswith(fill_prefix):
            style_list[i] = fill_prefix + color_string
            found_fill = True
            break
    if not found_fill:
        style_list.append(fill_prefix + color_string)
    return ';'.join(style_list)


def edit_daily(date, outfile):
    logging.debug('Editing day %s' % str(date.date()))
    global DAILY_TEMPLATE
    if not DAILY_TEMPLATE:
        DAILY_TEMPLATE = xml.dom.minidom.parse(DAILY_TEMPLATE_FILE)
    collection = DAILY_TEMPLATE.documentElement
    if date.strftime('%a').startswith('S'):
        color = WEEKEND_COLOR
    else:
        color = WEEKDAY_COLOR
    for e in collection.getElementsByTagName('text'):
        if e.hasAttribute('inkscape:label'):
            if e.getAttribute('inkscape:label') == 'Day number':
                # set the text
                tspan = e.getElementsByTagName('tspan')[0]
                tspan.childNodes[0].data = str(date.day)
                # set the color
                e.setAttribute('style', replace_fill_color(e.getAttribute('style'), color))
                tspan.setAttribute('style', replace_fill_color(e.getAttribute('style'), color))
            elif e.getAttribute('inkscape:label') == 'Day name':
                # set the text
                tspan = e.getElementsByTagName('tspan')[0]
                day_name = date.strftime('%a') # abbreviated, full is %A
                if day_name == 'Thu':
                    tspan.childNodes[0].data = 'R'
                elif day_name == 'Sun':
                    tspan.childNodes[0].data = 'O'
                else:
                    tspan.childNodes[0].data = day_name[0]
                # set the color
                e.setAttribute('style', replace_fill_color(e.getAttribute('style'), color))
                tspan.setAttribute('style', replace_fill_color(e.getAttribute('style'), color))
            elif e.getAttribute('inkscape:label') == 'Month':
                e.getElementsByTagName('tspan')[0].childNodes[0].data = date.strftime('%B')
            elif e.getAttribute('inkscape:label') == 'Year':
                e.getElementsByTagName('tspan')[0].childNodes[0].data = str(date.year)
    with open(outfile, 'w') as f:
        DAILY_TEMPLATE.writexml(f)
    return date.strftime('%b %d')


def edit_weekly(date, outfile):
    logging.debug('Editing week of %s' % str(date.date()))
    global WEEKLY_TEMPLATE
    if not WEEKLY_TEMPLATE:
        WEEKLY_TEMPLATE = xml.dom.minidom.parse(WEEKLY_TEMPLATE_FILE)
    collection = WEEKLY_TEMPLATE.documentElement
    week_start = date
    while week_start.strftime('%a') != 'Mon':
        week_start -= timedelta(days=1)
    week_end = week_start+timedelta(days=6)
    for e in collection.getElementsByTagName('text'):
        if e.hasAttribute('inkscape:label'):
            if e.getAttribute('inkscape:label') == 'Monday Date':
                e.getElementsByTagName('tspan')[0].childNodes[0].data = str(week_start.day)
            if e.getAttribute('inkscape:label') == 'Tuesday Date':
                e.getElementsByTagName('tspan')[0].childNodes[0].data = str((week_start+timedelta(days=1)).day)
            if e.getAttribute('inkscape:label') == 'Wednesday Date':
                e.getElementsByTagName('tspan')[0].childNodes[0].data = str((week_start+timedelta(days=2)).day)
            if e.getAttribute('inkscape:label') == 'Thursday Date':
                e.getElementsByTagName('tspan')[0].childNodes[0].data = str((week_start+timedelta(days=3)).day)
            if e.getAttribute('inkscape:label') == 'Friday Date':
                e.getElementsByTagName('tspan')[0].childNodes[0].data = str((week_start+timedelta(days=4)).day)
            if e.getAttribute('inkscape:label') == 'Saturday Date':
                e.getElementsByTagName('tspan')[0].childNodes[0].data = str((week_start+timedelta(days=5)).day)
            if e.getAttribute('inkscape:label') == 'Sunday Date':
                e.getElementsByTagName('tspan')[0].childNodes[0].data = str((week_start+timedelta(days=6)).day)
            elif e.getAttribute('inkscape:label') == 'Month':
                if week_start.month == (week_start+timedelta(days=6)).month:
                    e.getElementsByTagName('tspan')[0].childNodes[0].data = week_start.strftime('%B')
                else:
                    months = week_start.strftime('%B') + ' - ' + (week_start + timedelta(days=6)).strftime('%B')
                    e.getElementsByTagName('tspan')[0].childNodes[0].data = months
            elif e.getAttribute('inkscape:label') == 'Year':
                e.getElementsByTagName('tspan')[0].childNodes[0].data = str(date.year)
    with open(outfile, 'w') as f:
        WEEKLY_TEMPLATE.writexml(f)
    bookmark = week_start.strftime('%b %d - ')
    if week_start.month == week_end.month:
        bookmark += week_end.strftime('%d')
    else:
        bookmark += week_end.strftime('%b %d')
    return bookmark


def make_planner(filename, start_date, end_date):
    global MERGER
    logging.info('Making planner %s from %s to %s' % (filename, str(start_date), str(end_date)))
    with tempfile.TemporaryDirectory() as tmpdirname:
        date = start_date
        temp_weekly_svg  = os.path.join(tmpdirname, 'weekly.svg')
        temp_daily_svg   = os.path.join(tmpdirname, 'daily.svg')
        temp_weekly_pdf  = os.path.join(tmpdirname, 'weekly.pdf')
        temp_daily_pdf   = os.path.join(tmpdirname, 'daily.pdf')
        week_bookmark = None
        while date <= end_date:
            if date.strftime('%A') == "Monday":
                week_bookmark_name = edit_weekly(date, temp_weekly_svg)
                svg2pdf(temp_weekly_svg, temp_weekly_pdf)
                week_bookmark = add_page_to_merger(temp_weekly_pdf, bookmark_name=week_bookmark_name)
            day_bookmark_name = edit_daily(date, temp_daily_svg)
            svg2pdf(temp_daily_svg, temp_daily_pdf)
            if week_bookmark:
                add_page_to_merger(temp_daily_pdf, bookmark_name=day_bookmark_name, bookmark_parent=week_bookmark)
            else:
                add_page_to_merger(temp_daily_pdf)
            date += timedelta(days=1)
        with open(filename, 'wb') as f:
            MERGER.write(f)


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(prog="Make Planner",
            description="Create a pdf planner off template in svg",
            epilog="May your life feel under control.")
    parser.add_argument('filename')
    parser.add_argument('start_date', nargs='?', type=lambda d: datetime.strptime(d, '%Y-%m-%d').date(), default=datetime.now(), help='YYYY-MM-DD')
    parser.add_argument('end_date', nargs='?', type=lambda d: datetime.strptime(d, '%Y-%m-%d').date(), default=datetime.now()+timedelta(days=31), help='YYYY-MM-DD')
    parser.add_argument('-a', '--alternate_dailies_file', help="JSON file with dict keys as str YYYYMMDD and values of template filenames")
    parser.add_argument('-v', '--verbose', action='store_true')
    args = parser.parse_args()
    if args.verbose:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(level=logging.INFO)
    if args.alternate_dailies_file:
        make_alterate_dailies_dict(args.alternate_dailies_file)    
    make_planner(filename=args.filename, start_date=args.start_date, end_date=args.end_date)




