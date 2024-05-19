import sys
import openpyxl
import time
import webbrowser
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import os
import traceback
import datetime
from time import sleep
from datetime import datetime, timedelta, time
import platform

def read_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = list(sheet.values)
    columns = data[0]
    data = data[1:]
    return columns, data

# Function to generate the list of times and corresponding clock times
def generate_time_intervals(min_time, max_time, first_time, first_clock):
    # Create a list to hold the time intervals and corresponding clock times
    time_intervals = []

    # Convert min_time, max_time, and first_time to datetime objects
    min_datetime = datetime.combine(datetime.today(), min_time)
    max_datetime = datetime.combine(datetime.today(), max_time)
    max_datetime = max_datetime + timedelta(minutes=10)
    print('generate_time_intervals: %s %s' % (min_datetime, type(min_datetime)), file=sys.stderr)
    print('generate_time_intervals: %s %s' % (max_datetime, type(max_datetime)), file=sys.stderr)
    first_datetime = datetime.combine(datetime.today(), first_time)

    # Parse the first_clock string into a datetime.time object
    first_clock_time = datetime.strptime(first_clock, "%H:%M:%S").time()
    first_clock_datetime = datetime.combine(datetime.today(), first_clock_time)

    # Calculate the actual first clock time
    actual_first_clock_datetime = first_clock_datetime - timedelta(
        hours=first_datetime.hour, minutes=first_datetime.minute, seconds=first_datetime.second)

    # Generate the time intervals and corresponding clock times
    current_time = min_datetime
    while current_time <= max_datetime:
        # Calculate the corresponding clock time
        clock_time = actual_first_clock_datetime + (current_time - min_datetime)
        # Append the tuple (current time, corresponding clock time) to the list
        time_intervals.append((current_time.time(), clock_time.time().strftime("%H:%M:%S")))
        current_time += timedelta(minutes=1)

    return time_intervals

def process_data(columns, data):
    #columns = ['Stopwatch', 'Bib', 'LastName', 'FirstName', 'Gender', 'Clock']

    data = [dict(zip(columns, row)) for row in data]

    # Convert Stopwatch to timedelta
    for i, row in enumerate(data):
        #print('process_data[%d] Bib: %s %s' % (i, row['Bib'], type(row['Bib'])), file=sys.stderr)
        if not row['Bib']:
            row['Bib'] = ''
        row['Bib'] = str(f"  {row['Bib']}")

    first_clock = data[0]['Clock']
    first_time = data[0]['Stopwatch']
    min_time = datetime.strptime('00:00:00', '%H:%M:%S').time()
    max_time = max(row['Stopwatch'] for row in data)

    time_intervals = generate_time_intervals(min_time, max_time, first_time, first_clock)

    #print('process_data: %s %s' % (first_clock, type(first_clock)), file=sys.stderr)
    #print('process_data: %s %s' % (first_time, type(first_time)), file=sys.stderr)
    #print('process_data: %s %s' % (min_time, type(min_time)), file=sys.stderr)
    #print('process_data: %s %s' % (max_time, type(max_time)), file=sys.stderr)
    #print('process_data: %s' % time_intervals, file=sys.stderr)

    # Normalize time range to start from 0
    #for row in data:

    # Generate full range of time in seconds
    #full_range = range(0, max_time - min_time + 1, 60)
    #full_data = [{'Stopwatch': t} for t in full_range]
    # Create a list to hold the datetime objects


    full_data = [{'Stopwatch': t, 'Clock': c, 'Start Time': c} for t, c in time_intervals]

    # Merge data with full range
    merged_data = []
    data_dict = {row['Stopwatch']: row for row in data}
    for row in full_data:
        if row['Stopwatch'] in data_dict:
            stop_time = row['Stopwatch']
            row = data_dict[row['Stopwatch']]
            row['Start Time'] = row['Clock']
            merged_data.append(row)
        else:
            empty_row = {col: '' for col in columns}
            empty_row['Stopwatch'] = row['Stopwatch']
            empty_row['Clock'] = row['Clock']
            empty_row['Start Time'] = row['Start Time']
            merged_data.append(empty_row)

    # Convert Stopwatch back to HH:MM:SS and add Clock column
    #for row in merged_data:
    #    t = row['Stopwatch']
    #    row['Stopwatch'] = f"{t // 3600:02}:{(t % 3600) // 60:02}:{t % 60:02}"
    #    row['Clock'] = (t + 10 * 3600) % 86400
    #    row['Clock'] = f"{row['Clock'] // 3600:02}:{(row['Clock'] % 3600) // 60:02}:{row['Clock'] % 60:02}"

    return columns, merged_data

def generate_pdf(columns, data, output_path, filename):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    margin = 0.5 * inch
    rows_per_page = 30
    row_height = (height - 2 * margin) / (rows_per_page + 1)
    
    header = [("Stopwatch", 0.6), 
              ("Bib", 0.5), 
              ("LastName", 1.), ("FirstName", 1), 
              ("Notes/Delay", 1), 
              ("Start Time", 0.6)]
    xlength = sum([header[i][1] for i in range(len(header))])
    xs = [0]
    x = 0
    for i in range(len(header)):
        x += header[i][1]
        xs.append(x)
    adjust = 7.5 / xs[-1]
    xs = [x * adjust for x in xs]

    def draw_page_header(xs, page):
        c.setFont("Helvetica-Bold", 8)

        c.drawString(margin, height-.2*inch, filename)
        c.drawString(margin+7*inch, height-.2*inch, f"Page {page}")
        c.drawString(margin+.2*inch, .55*inch, 'Notes: ___________________________________________________________________________________________________')
        c.drawString(margin+.2*inch, .35*inch, '_________________________________________________________________________________________________________')
        c.drawString(margin+.2*inch, .1*inch, 'Starter: _________________________________________________________________   Actual Start Time: _______________')

        for i, (name, width) in enumerate(header):
            x = margin + xs[i] * inch
            c.drawString(x, height - margin, name)
        for i in range(len(xs)):
            x = margin + xs[i] * inch
            c.line(x, height - margin - 2, x, 1.62*margin)
        c.line(margin, height - margin - 2, 8*inch, height - margin - 2)

    def draw_row(xs, index, row):
        y = height - margin - (index + 1) * row_height
        c.setFont("Helvetica", 16)
        row['Gender'] = ' M ' if row['Gender'] == 'Men' else ' F '
        
        for i, (name, width) in enumerate(header):
            x = margin + xs[i] * inch
            value = row.get(name, '')
            c.drawString(x + 2, y + row_height / 2 - .1*inch, str(value))
        c.line(margin, y, 8*inch, y)
    
    page = 1
    for i, row in enumerate(data):
        if i % rows_per_page == 0:
            if i > 0:
                c.showPage()
            draw_page_header(xs, page)
            page += 1
        draw_row(xs, i % rows_per_page, row)
    
    c.save()

def main():
    if len(sys.argv) < 2:
        print("Usage: python tt2pdf.py <xlsx_file>")
        sys.exit(1)

    xlsx_file = sys.argv[1]
    pdf_file = xlsx_file.replace(".xlsx", ".pdf")
    filename = os.path.basename(xlsx_file).replace(".xlsx", "")
    if not os.path.isfile(xlsx_file):
        print(f"File {xlsx_file} does not exist.")
        sys.exit(1)

    try:
        columns, data = read_excel(xlsx_file)
        columns, processed_data = process_data(columns, data)
        generate_pdf(columns, processed_data, pdf_file, filename)
    except Exception as e:
        print(f"An error occurred: {e}")
        print(traceback.format_exc())
        sys.exit(1)

    if platform.system() == 'Windows':
        print('pdf_file: %s' % (pdf_file), file=sys.stderr)
        print('filename: %s' % (filename), file=sys.stderr)
        if pdf_file.startswith(filename):
            url = f"file:///{os.getcwd()}/{pdf_file}"
        else:
            url = f"file:///{pdf_file}"
        #3cwd = os.getcwd()
        #print('cwd: %s' % (cwd), file=sys.stderr)
        webbrowser.open_new_tab(url)
        print('url: %s' % (url), file=sys.stderr)
        #sleep(120)
    #else:
    #    url = f"file://{pdf_file}"
    #    webbrowser.open_new_tab(url)
    #    time.sleep(120)

if __name__ == "__main__":
    main()

