from PIL import Image
import os,io
import pyodbc

conn = pyodbc.connect('DRIVER=FreeTDS;SERVER=192.168.0.30;PORT=1433;DATABASE=007;UID=winkel;PWD=exactwinkel')
cursor = conn.cursor()
for row in cursor.execute('select TOP 5 id, picture from Items where picture is not null'):
    print(row[0])
    image = Image.open(io.BytesIO(row[1]))
    image.save('/home/exact/Downloads/pictures/{0}.jpg'.format(row.id))

conn.close()
