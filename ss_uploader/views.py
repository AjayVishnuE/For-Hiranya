import os
import tempfile
import pandas as pd                                                                                        # type: ignore
from io import BytesIO
from django.conf import settings
from reportlab.lib import colors                                                                           # type: ignore
from .forms import UploadFileForm
from reportlab.lib.units import cm                                                                         # type: ignore
from django.http import HttpResponse
from tempfile import NamedTemporaryFile
from reportlab.lib.pagesizes import landscape, A4                                                          # type: ignore
from datetime import datetime, timedelta
from reportlab.lib.enums import TA_CENTER, TA_LEFT                                                         # type: ignore
from django.shortcuts import render, redirect                                                              # type: ignore
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle                                       # type: ignore
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak          # type: ignore


# def upload_file(request):
#     if request.method == 'POST':
#         form = UploadFileForm(request.POST, request.FILES)
#         if form.is_valid():
#             file = request.FILES['file']
#             df = pd.read_excel(file)

#             # Store the dataframe in session for future processing (sorting, filtering)
#             request.session['df'] = df.to_dict(orient='records')  # Store data in session
#             columns = list(df.columns)  # Store column names dynamically
#             request.session['columns'] = columns  # Store the column names in session
#             return render(request, 'uploader/display_data.html', {'data': df.to_dict(orient='records'), 'columns': columns})

#     if request.method == 'GET':
#         # Check if the session data is available
#         df_data = request.session.get('df', None)
#         columns = request.session.get('columns', None)  # Retrieve column names from session

#         # If no session data, render the upload form instead of redirecting
#         if df_data is None or columns is None:
#             form = UploadFileForm()
#             return render(request, 'uploader/upload.html', {'form': form})

#         # If session data is available, display it
#         df = pd.DataFrame(df_data)

#         # Check if the "Clear Repeating Entries" button was clicked
#         clear_repeat = request.GET.get('clear_repeat', None)
#         if clear_repeat == 'clear_repetitions':
#             # Drop duplicates based on Name and Phone number
#             if 'NAME' in df.columns and 'PHONE' in df.columns:
#                 df = df.drop_duplicates(subset=['NAME', 'PHONE'])

                
#         # Get filter parameters dynamically for each column
#         filters = {
#             'NAME': request.GET.get('name_filter', None),
#             'RMS': request.GET.get('rms_filter', None),
#             'PIN': request.GET.get('pin_filter', None)
#         }

#         # Apply filters based on filter input
#         for col, filter_value in filters.items():
#             if filter_value:
#                 df = df[df[col].astype(str).str.contains(filter_value, case=False, na=False)]

#         # After filtering, apply the sorting logic
#         if 'RMS' in columns and 'PIN' in columns and 'NAME' in columns:
#             df = df.sort_values(by=['RMS', 'PIN', 'NAME'], ascending=[True, True, True])

#         # Drop duplicate rows based on all columns (or specific columns if needed)
#         df = df.drop_duplicates()

#         # Store the filtered, sorted, and deduplicated data back into the session
#         request.session['filtered_df'] = df.to_dict(orient='records')

#         # Render the sorted, filtered, and deduplicated data to the template
#         return render(request, 'uploader/display_data.html', {'data': df.to_dict(orient='records'), 'columns': columns})

#     # Render the upload form in case of any other requests
#     form = UploadFileForm()
#     return render(request, 'uploader/upload.html', {'form': form})

def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            file = request.FILES['file']
            df = pd.read_excel(file)

            # Save the DataFrame to two temporary CSV files
            # One for original data (never changes)
            temp_file_original = tempfile.NamedTemporaryFile(delete=False, dir=settings.MEDIA_ROOT, suffix='.csv')
            df.to_csv(temp_file_original.name, index=False)

            # Another for filtered data (this gets updated with filters/sorts)
            temp_file_filtered = tempfile.NamedTemporaryFile(delete=False, dir=settings.MEDIA_ROOT, suffix='.csv')
            df.to_csv(temp_file_filtered.name, index=False)

            # Store the file paths and column names in the session
            request.session['temp_file_path_original'] = temp_file_original.name  # Original unmodified data
            request.session['temp_file_path_filtered'] = temp_file_filtered.name  # Filtered data
            request.session['columns'] = list(df.columns)  # Store only column names in session

            return render(request, 'uploader/display_data.html', {
                'data': df.to_dict(orient='records'), 
                'columns': list(df.columns)
            })

    if request.method == 'GET':
        # Retrieve the file paths and column names from the session
        temp_file_path_original = request.session.get('temp_file_path_original', None)
        temp_file_path_filtered = request.session.get('temp_file_path_filtered', None)
        columns = request.session.get('columns', None)

        if temp_file_path_original and columns:
            # Load the original data first (to apply filters on the original data)
            df = pd.read_csv(temp_file_path_original)

            # Apply filters dynamically for each column
            filters = {
                'NAME': request.GET.get('name_filter', None),
                'RMS': request.GET.get('rms_filter', None),
                'PIN': request.GET.get('pin_filter', None)
            }

            # Apply filters based on filter input
            for col, filter_value in filters.items():
                if filter_value:
                    df = df[df[col].astype(str).str.contains(filter_value, case=False, na=False)]

            # Apply sorting logic (sort by 'RMS', 'PIN', and 'NAME')
            if 'RMS' in columns and 'PIN' in columns and 'NAME' in columns:
                df = df.sort_values(by=['RMS', 'PIN', 'NAME'], ascending=[True, True, True])

            # Drop duplicate rows based on all columns
            df = df.drop_duplicates()

            # Save the filtered and sorted data back to the filtered CSV file
            df.to_csv(temp_file_path_filtered, index=False)

            # Render the sorted, filtered, and deduplicated data
            return render(request, 'uploader/display_data.html', {
                'data': df.to_dict(orient='records'), 
                'columns': columns
            })

    # If no POST or GET request, return the upload form
    form = UploadFileForm()
    return render(request, 'uploader/upload.html', {'form': form})






# def clear_file(request):
#     # Clear the session data
#     if 'df' in request.session:
#         del request.session['df']
#     if 'columns' in request.session:
#         del request.session['columns']
#     if 'filtered_df' in request.session:
#         del request.session['filtered_df']
    
#     # Redirect back to the file upload page
#     return redirect('upload_file')
def clear_file(request):
    # Get both temp file paths from session
    temp_file_path_original = request.session.get('temp_file_path_original', None)
    temp_file_path_filtered = request.session.get('temp_file_path_filtered', None)

    # Remove the original temp file if it exists
    if temp_file_path_original and os.path.exists(temp_file_path_original):
        os.remove(temp_file_path_original)

    # Remove the filtered temp file if it exists
    if temp_file_path_filtered and os.path.exists(temp_file_path_filtered):
        os.remove(temp_file_path_filtered)

    # Clear the session data for temp files and columns
    if 'temp_file_path_original' in request.session:
        del request.session['temp_file_path_original']
    if 'temp_file_path_filtered' in request.session:
        del request.session['temp_file_path_filtered']
    if 'columns' in request.session:
        del request.session['columns']

    return redirect('upload_file')


# To create a pdf of expired, expiring and to expire customers

def download_expired_customers(request):
    # Retrieve the temporary CSV file path from session
    temp_file_path = request.session.get('temp_file_path_original', None)
    
    if not temp_file_path or not os.path.exists(temp_file_path):
        return HttpResponse("No data available for Excel download.")
    
    # Load the data from the temporary CSV file
    df = pd.read_csv(temp_file_path)

    # Columns to include in the Excel file
    columns = ['NAME', 'ADDRESS', 'LOCATION', 'POST', 'DISTRICT', 'PHONE', 'STATE',
               'FROM DATE', 'DURATION(M)', 'CLOSING DATE', 'STATUS', 'INTRODUCER NAME',
               'INTRODUCER VEDAVAHINI', 'INTRODUCER PHONE NO.']

    # Convert 'CLOSING DATE' to datetime for comparison
    df['CLOSING DATE'] = pd.to_datetime(df['CLOSING DATE'], format='%d-%m-%Y')

    # Get current date and calculate previous, current, and next months
    today = datetime.today()
    start_of_current_month = today.replace(day=1)  # Beginning of the current month
    start_of_next_month = (start_of_current_month + timedelta(days=32)).replace(day=1)  # Beginning of next month
    start_of_previous_month = (start_of_current_month - timedelta(days=1)).replace(day=1)  # Beginning of the previous month

    # Filter the data where the 'CLOSING DATE' is within the previous, current, or next month
    df_filtered = df[
        (df['CLOSING DATE'] >= start_of_previous_month) & 
        (df['CLOSING DATE'] < start_of_next_month + timedelta(days=32))
    ]

    # Sort by 'CLOSING DATE', 'INTRODUCER VEDAVAHINI', and 'NAME'
    df_filtered = df_filtered.sort_values(by=['CLOSING DATE', 'INTRODUCER VEDAVAHINI', 'NAME'])

    # Convert 'CLOSING DATE' back to 'dd-mm-yyyy' format
    df_filtered['CLOSING DATE'] = df_filtered['CLOSING DATE'].dt.strftime('%d-%m-%Y')

    # Create a BytesIO buffer to hold the Excel data
    buffer = BytesIO()

    # Write the filtered DataFrame to the buffer using XlsxWriter as the engine
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_filtered.to_excel(writer, sheet_name='Expired Customers', index=False, columns=columns)

    # Set up the HttpResponse to download the file as an Excel attachment
    response = HttpResponse(buffer.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="expired_customers.xlsx"'

    return response
# def download_expired_customers(request):
#     # Load the data from session
#     df_data = request.session.get('df', None)
    
#     if not df_data:
#         return HttpResponse("No data available for Excel download.")
    
#     df = pd.DataFrame(df_data)

#     # Columns to include in the Excel file
#     columns = ['NAME', 'ADDRESS', 'LOCATION', 'POST', 'DISTRICT', 'PHONE', 'STATE',
#                'FROM DATE', 'DURATION(M)', 'CLOSING DATE', 'STATUS', 'INTRODUCER NAME',
#                'INTRODUCER VEDAVAHINI', 'INTRODUCER PHONE NO.']

#     # Convert 'CLOSING DATE' to datetime for comparison
#     df['CLOSING DATE'] = pd.to_datetime(df['CLOSING DATE'], format='%d-%m-%Y')

#     # Get current date and calculate previous, current, and next months
#     today = datetime.today()
#     start_of_current_month = today.replace(day=1)  # Beginning of the current month
#     start_of_next_month = (start_of_current_month + timedelta(days=32)).replace(day=1)  # Beginning of next month
#     start_of_previous_month = (start_of_current_month - timedelta(days=1)).replace(day=1)  # Beginning of the previous month

#     # Filter the data where the 'CLOSING DATE' is within the previous, current, or next month
#     df_filtered = df[
#         (df['CLOSING DATE'] >= start_of_previous_month) & 
#         (df['CLOSING DATE'] < start_of_next_month + timedelta(days=32))
#     ]

#     # Sort by 'CLOSING DATE', 'INTRODUCER VEDAVAHINI', and 'NAME'
#     df_filtered = df_filtered.sort_values(by=['CLOSING DATE', 'INTRODUCER VEDAVAHINI', 'NAME'])

#     # Convert 'CLOSING DATE' back to 'dd-mm-yyyy' format
#     df_filtered['CLOSING DATE'] = df_filtered['CLOSING DATE'].dt.strftime('%d-%m-%Y')

#     # Create a BytesIO buffer to hold the Excel data
#     buffer = BytesIO()

#     # Write the DataFrame to the buffer using XlsxWriter as the engine
#     with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
#         df_filtered.to_excel(writer, sheet_name='Expired Customers', index=False, columns=columns)

#     # Set up the HttpResponse to download the file as an Excel attachment
#     response = HttpResponse(buffer.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
#     response['Content-Disposition'] = 'attachment; filename="expired_customers.xlsx"'

#     return response

def convert_to_pdf(request):
    # Load the data from the temporary CSV file (replace with your actual temporary file path)
    temp_file_path = request.session.get('temp_file_path_filtered')
    
    if not temp_file_path:
        return HttpResponse("No data available for PDF conversion.")
    
    try:
        df = pd.read_csv(temp_file_path)
    except Exception as e:
        return HttpResponse(f"Error reading the CSV file: {str(e)}")
    
    # Create a response object for the PDF
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="customer_labels.pdf"'

    # Buffer to hold PDF data
    buffer = BytesIO()

    # Create a custom PDF document with the given size (45 cm width, 14 cm height)
    pdf_width = 45 * cm
    pdf_height = 14 * cm
    pdf = SimpleDocTemplate(buffer, pagesize=(pdf_width, pdf_height), leftMargin=1 * cm, rightMargin=1 * cm)

    # Define the styles for the text
    styles = getSampleStyleSheet()
    style_bold_center = ParagraphStyle(
        'BoldCenter', 
        parent=styles['Normal'], 
        fontName='Helvetica-Bold', 
        alignment=TA_CENTER
    )
    style_left = ParagraphStyle('left', alignment=TA_LEFT)
    elements = []  # List to store elements for PDF generation

    # Static "From" address
    from_address =  """
                    <b>FROM</b><br/><br/>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>HIRANYA MAGAZINE</b><br/>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;12/2375 C.N. ARCADE<br/>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Florican Road<br/>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Malaparamba, Kozhikode<br/>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Kerala – 673 009<br/>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ph: 0495-2961151
                    """

    # Loop through the filtered data and create one page per "To" address
    for index, row in df.iterrows():
        # Add the "BOOK POST" title for each entry
        elements.append(Spacer(1, 0.5 * cm))  # Spacer above the title
        elements.append(Paragraph("BOOK POST", style_bold_center))
        elements.append(Spacer(1, 0.5 * cm))  # Spacer below the title

        # "To" address details from the row
        to_address = f"""
        <b>TO</b><br/><br/>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>{row['NAME']}</b><br/>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['ADDRESS']}<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['PLACE']}<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['LOCATION']}<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['POST']}<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['DISTRICT']}<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['STATE']}<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['PIN']}<br/>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['RMS']}<br/>
        """

        # Create paragraphs for "From" and "To" addresses
        from_address_paragraph = Paragraph(from_address, style_left)
        to_address_paragraph = Paragraph(to_address, style_left)

        # Create a table to structure the "From" and "To" addresses on the page
        data = [
            [" ", from_address_paragraph, to_address_paragraph, " "],  # Left: From, Right: To
        ]

        # Create the table with adjusted column widths and left alignment for addresses
        table = Table(data, colWidths=[pdf_width/3, pdf_width/6, pdf_width/6, pdf_width/3])  # Add a small column for spacing
        table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('LEFTPADDING', (0, 0), (-1, -1), 50),  # Add padding to move the text away from the border
        ]))

        # Add the table to the elements list
        elements.append(table)

        # Add a page break after each entry to ensure it starts on a new page
        elements.append(PageBreak())

    # Build the PDF
    pdf.build(elements)

    # Get the PDF value from the buffer and return it as a response
    response.write(buffer.getvalue())
    buffer.close()

    return response

# def convert_to_pdf(request):
#     # Fetch the filtered data from the session (this should be the data after applying filters)
#     df_data = request.session.get('filtered_df', None)

#     if not df_data:
#         return HttpResponse("No data available for PDF conversion.")

#     df = pd.DataFrame(df_data)

#     # Create a response object for the PDF
#     response = HttpResponse(content_type='application/pdf')
#     response['Content-Disposition'] = 'attachment; filename="customer_labels.pdf"'

#     # Buffer to hold PDF data
#     buffer = BytesIO()

#     # Create a custom PDF document with the given size (45 cm width, 14 cm height)
#     pdf_width = 45 * cm
#     pdf_height = 14 * cm
#     pdf = SimpleDocTemplate(buffer, pagesize=(pdf_width, pdf_height), leftMargin=1 * cm, rightMargin=1 * cm)

#     # Define the styles for the text
#     styles = getSampleStyleSheet()
#     style_bold_center = ParagraphStyle(
#         'BoldCenter', 
#         parent=styles['Normal'], 
#         fontName='Helvetica-Bold', 
#         alignment=TA_CENTER
#     )
#     style_left = ParagraphStyle('left', alignment=TA_LEFT)
#     elements = []  # List to store elements for PDF generation

#     # Static "From" address
#     from_address =  """
#                     <b>FROM</b><br/><br/>
#                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>HIRANYA MAGAZINE</b><br/>
#                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;12/2375 C.N. ARCADE<br/>
#                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Florican Road<br/>
#                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Malaparamba, Kozhikode<br/>
#                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Kerala – 673 009<br/>
#                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ph: 0495-2961151
#                     """

#     # Loop through the filtered data and create one page per "To" address
#     for index, row in df.iterrows():
#         # Add the "BOOK POST" title for each entry
#         elements.append(Spacer(1, 0.5 * cm))  # Spacer above the title
#         elements.append(Paragraph("BOOK POST", style_bold_center))
#         elements.append(Spacer(1, 0.5 * cm))  # Spacer below the title

#         # "To" address details from the row
#         to_address = f"""
#         <b>TO</b><br/><br/>
#         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>{row['NAME']}</b><br/>
#         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['ADDRESS']}<br/>
#         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['PLACE']}<br/>
#         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['LOCATION']}<br/>
#         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['POST']}<br/>
#         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['DISTRICT']}<br/>
#         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['STATE']}<br/>
#         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['PIN']}<br/>
#         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{row['RMS']}<br/>
#         """

#         # Create paragraphs for "From" and "To" addresses
#         from_address_paragraph = Paragraph(from_address, style_left)
#         to_address_paragraph = Paragraph(to_address, style_left)

#         # Create a table to structure the "From" and "To" addresses on the page
#         data = [
#             [" ", from_address_paragraph, to_address_paragraph, " "],  # Left: From, Right: To
#         ]

#         # Create the table with adjusted column widths and left alignment for addresses
#         table = Table(data, colWidths=[pdf_width/3, pdf_width/6, pdf_width/6, pdf_width/3])  # Add a small column for spacing
#         table.setStyle(TableStyle([
#             ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
#             ('VALIGN', (0, 0), (-1, -1), 'TOP'),
#             ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
#             ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
#             ('FONTSIZE', (0, 0), (-1, -1), 12),
#             ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
#             ('LEFTPADDING', (0, 0), (-1, -1), 50),  # Add padding to move the text away from the border
#         ]))

#         # Add the table to the elements list
#         elements.append(table)

#         # Add a page break after each entry to ensure it starts on a new page
#         elements.append(PageBreak())

#     # Build the PDF
#     pdf.build(elements)

#     # Get the PDF value from the buffer and return it as a response
#     response.write(buffer.getvalue())
#     buffer.close()

#     return response


