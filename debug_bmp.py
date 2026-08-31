import struct, os, tempfile

tmp = os.path.join(tempfile.gettempdir(), 'palette_sample.bmp')
exists = os.path.exists(tmp)
size = os.path.getsize(tmp) if exists else 0
print('File exists:', exists)
print('File size:', size, 'bytes')

if not exists:
    print('File not found - PS may not have created it')
    exit()

with open(tmp, 'rb') as f:
    data = f.read()

print('Magic:', data[:2])
print('Data offset:', struct.unpack('<I', data[10:14])[0])
print('Header size:', struct.unpack('<I', data[14:18])[0])
print('Width:', struct.unpack('<i', data[18:22])[0])
print('Height:', struct.unpack('<i', data[22:26])[0])
print('BPP:', struct.unpack('<H', data[28:30])[0])
print('Compression:', struct.unpack('<I', data[30:34])[0])
print('First 60 bytes hex:', data[:60].hex())
