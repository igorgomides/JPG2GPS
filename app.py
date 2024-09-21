from flask import Flask, request, render_template, send_file
from PIL import Image
import piexif
import os
import pandas as pd
import xlsxwriter
import requests
import tempfile

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

def get_decimal_from_dms(dms, ref):
    degrees = dms[0][0] / dms[0][1]
    minutes = dms[1][0] / dms[1][1] / 60.0
    seconds = dms[2][0] / dms[2][1] / 3600.0
    decimal = degrees + minutes + seconds
    if ref in ['S', 'W']:
        decimal = -decimal
    return decimal

def get_gps_location(exif_data):
    gps_info = exif_data.get("GPS")
    if gps_info:
        gps_latitude = gps_info[2]
        gps_latitude_ref = gps_info[1].decode()
        gps_longitude = gps_info[4]
        gps_longitude_ref = gps_info[3].decode()

        lat = get_decimal_from_dms(gps_latitude, gps_latitude_ref)
        lon = get_decimal_from_dms(gps_longitude, gps_longitude_ref)
        return lat, lon
    return None

def get_address(lat, lon, api_key):
    url = f"https://maps.googleapis.com/maps/api/geocode/json?latlng={lat},{lon}&key={api_key}"
    response = requests.get(url)
    if response.status_code == 200:
        results = response.json().get('results')
        if results:
            return results[0].get('formatted_address')
    return None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    files = request.files.getlist('photos')
    api_key = request.form['api_key']
    data = []

    with tempfile.TemporaryDirectory() as temp_dir:
        for file in files:
            filename = os.path.join(temp_dir, file.filename)
            file.save(filename)
            image = Image.open(filename)
            exif_data = piexif.load(image.info.get('exif', b''))
            
            location = get_gps_location(exif_data)
            if location:
                lat, lon = location
                address = get_address(lat, lon, api_key)
                data.append({
                    "File": filename,
                    "Filename": file.filename,
                    "Latitude": lat,
                    "Longitude": lon,
                    "Address": address
                })
            else:
                data.append({
                    "File": filename,
                    "Filename": file.filename,
                    "Latitude": None,
                    "Longitude": None,
                    "Address": None
                })

        df = pd.DataFrame(data)
        output_path = os.path.join(temp_dir, 'gps_data.xlsx')
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='GPS Data', startrow=1, header=False)
            
            workbook  = writer.book
            worksheet = writer.sheets['GPS Data']
            
            header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1})
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            for row_num, row_data in df.iterrows():
                worksheet.write_url(row_num + 1, 0, f'external:{row_data["File"]}', string=row_data["Filename"])
                worksheet.insert_image(row_num + 1, len(df.columns), row_data["File"], {'x_scale': 0.1, 'y_scale': 0.1})

        return send_file(output_path, as_attachment=True, download_name='gps_data.xlsx')

if __name__ == '__main__':
    app.run(debug=True)