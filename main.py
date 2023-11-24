import pandas as pd
import xml.etree.ElementTree as ET

# Excel dosyasının yolu
excel_file_path = './All_Coordinates.xlsx'

# Excel dosyasını oku
df = pd.read_excel(excel_file_path)

# Boş enlem veya boylam değerlerini filtrele
df = df[df['lng'].notna() & df['lat'].notna()]

# KML yapısını oluştur
kml = ET.Element('kml', xmlns="http://www.opengis.net/kml/2.2")
document = ET.SubElement(kml, 'Document')

# Document içindeki name elementini değiştir
doc_name = ET.SubElement(document, 'name')
doc_name.text = "tibb_temp"  # Burada Excel'den alınan bir değeri kullanabilirsiniz.


for _, row in df.iterrows():
    # NaN değerler için boş string veya varsayılan bir değer kullan
    description_parts = [
        f"Ad: {row['AD']}" if pd.notna(row['AD']) else "",
        f"Region: {row['region']}" if pd.notna(row['region']) else "",
        f"Tip: {row['muesisse_tipi']}" if pd.notna(row['muesisse_tipi']) else "",
        f"Ünvan: {row['unvan']}" if pd.notna(row['unvan']) else "",
        f"Tel: {row['Tel']}" if pd.notna(row['Tel']) else "",
        f"Rəhbər: {row['Rehberin_adi']}" if pd.notna(row['Rehberin_adi']) else "",
        f"Növ: {row['nov']}" if pd.notna(row['nov']) else "",
        # Diğer alanlar için benzer kontrol
    ]
    description = "\n".join(part for part in description_parts if part)

    placemark = ET.SubElement(document, 'Placemark')
    ET.SubElement(placemark, 'name').text = str(row['AD'])
    desc_st = ET.SubElement(placemark, 'description')
    desc_st.set("color", "99307b19")  # Attribute ekleme
    desc_st.set("width", "80.0")  # Attribute ekleme
    ET.SubElement(placemark, 'description').text = description
    
    point = ET.SubElement(placemark, 'Point')
    ET.SubElement(point, 'coordinates').text = f"{row['lng']},{row['lat']},0"

# KML dosyasını kaydet
output_kml_file = './test.kml'
tree = ET.ElementTree(kml)
tree.write(output_kml_file, encoding='utf-8', xml_declaration=True)
