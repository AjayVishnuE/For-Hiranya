import pandas as pd
from io import BytesIO
from reportlab.lib import colors
from .forms import UploadFileForm
from reportlab.lib.units import cm
from django.http import HttpResponse
from reportlab.lib.pagesizes import A4
from django.shortcuts import render, redirect
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer

def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            file = request.FILES['file']
            df = pd.read_excel(file)

            # Store the dataframe in session for future processing (sorting, filtering)
            request.session['df'] = df.to_dict(orient='records')
            columns = df.columns
            return render(request, 'uploader/display_data.html', {'data': df.to_dict(orient='records'), 'columns': columns})

    if request.method == 'GET':
        df_data = request.session.get('df', None)

        if df_data:
            df = pd.DataFrame(df_data)

            # Log the available columns to check if 'Pin' exists
            available_columns = df.columns
            name_exists = 'Name' in available_columns
            address_exists = 'Address' in available_columns
            pin_exists = 'Pin' in available_columns

            # Get filter parameters and check for column existence before filtering
            name_filter = request.GET.get('name_filter', None)
            if name_filter and name_exists:
                df = df[df['Name'].str.contains(name_filter, case=False, na=False)]

            address_filter = request.GET.get('address_filter', None)
            if address_filter and address_exists:
                df = df[df['Address'].str.contains(address_filter, case=False, na=False)]

            pin_filter = request.GET.get('pin_filter', None)
            if pin_filter and pin_exists:
                df['Pin'] = df['Pin'].astype(str)  # Ensure Pin is treated as a string
                df = df[df['Pin'].str.contains(pin_filter, case=False, na=False)]  # Filter by Pin

            # Sort the DataFrame first by Name, then by Pin
            if name_exists and pin_exists:
                df = df.sort_values(by=['Name', 'Pin'], ascending=[True, True])

            sort_by = request.GET.get('sort_by', None)
            if sort_by == 'name_pin':
                # Sort by Name first, then Pin
                df = df.sort_values(by=['Name', 'Pin'], ascending=[True, True])
            elif sort_by == 'pin_name':
                # Sort by Pin first, then Name
                df = df.sort_values(by=['Pin', 'Name'], ascending=[True, True])

            # Store the filtered data in session for later use in PDF generation
            request.session['filtered_df'] = df.to_dict(orient='records')

            columns = df.columns
            return render(request, 'uploader/display_data.html', {'data': df.to_dict(orient='records'), 'columns': columns})

    form = UploadFileForm()
    return render(request, 'uploader/upload.html', {'form': form})


def clear_file(request):
    # Clear the session data
    if 'df' in request.session:
        del request.session['df']
    
    # Redirect back to the file upload page
    return redirect('upload_file')


def convert_to_pdf(request):
    # Fetch the filtered data from session (this should be the data after applying filters)
    df_data = request.session.get('filtered_df', None)  # 'filtered_df' stores filtered data
    
    if not df_data:
        return HttpResponse("No filtered data available for PDF conversion.")

    # Create a response object for the PDF
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="filtered_data.pdf"'

    # Buffer to hold PDF data
    buffer = BytesIO()

    # Create a PDF object using SimpleDocTemplate with 0.3 cm margins
    pdf = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=0.3 * cm, rightMargin=0.3 * cm,
                            topMargin=0.2 * cm, bottomMargin=0.2 * cm)

    # Create the content for the PDF
    elements = []

    # Data table settings
    data = []
    row_counter = 0
    page_rows = []  # To store rows for each page (10 rows with 5 columns)

    # Calculate fixed column width and row height
    column_width = (21.0 - 2 * 0.3) / 5 * cm  # 4.08 cm width per column
    row_height = (29.0 - 2 * 0.3) / 10 * cm  # 2.91 cm height per row

    for row in df_data:
        # Extract the Name, Address, and Pin
        name = row.get('Name', '')
        address = row.get('Address', '')
        pin = row.get('Pin', '')

        # Format the data as "Name, Address, Pin" in vertical format
        entry = f"{name}\n{address}\n{pin}"

        # Add this entry to the current row (which will have 5 columns of data)
        page_rows.append(entry)

        # After every 5 entries, create a new row with 5 columns
        if len(page_rows) == 5:
            data.append(page_rows)
            page_rows = []  # Reset for the next row
            row_counter += 1

        # Every 10 rows (5x10 = 50 entries), add the table and reset
        if row_counter == 10:
            table = Table(data, colWidths=[column_width] * 5, rowHeights=[row_height] * 10)
            table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0, colors.transparent),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Align content to the top of the cell
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ('LEFTPADDING', (0, 0), (-1, -1), 15),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 10),
            ]))
            elements.append(table)
            data = []  # Reset data for the next page
            row_counter = 0

    # Pad remaining rows with empty strings to ensure exactly 10 rows
    if page_rows:
        while len(page_rows) < 5:
            page_rows.append('')  # Fill empty columns
        data.append(page_rows)

    if len(data) > 0:
        # Ensure that the last table has exactly 10 rows by padding with empty rows
        while len(data) < 10:
            data.append([''] * 5)  # Add empty rows

        # Create the final table for the remaining data
        table = Table(data, colWidths=[column_width] * 5, rowHeights=[row_height] * 10)
        table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0, colors.transparent),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING', (0, 0), (-1, -1), 15),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
        ]))
        elements.append(table)

    # Build the PDF
    pdf.build(elements)

    # Get the PDF value from the buffer
    response.write(buffer.getvalue())
    buffer.close()

    return response