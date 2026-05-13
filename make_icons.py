#!/usr/bin/env python3
"""
Generates the four PNG icons required by the Outlook manifest.
Run once: python3 make_icons.py
"""
import struct, zlib, os

def png(size, bg=(14,165,233), fg=(255,255,255)):
    """Generate a minimal valid PNG with a bell icon at given size."""
    def chunk(tag, data):
        c = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', c)

    # Build pixel grid
    pixels = []
    cx, cy = size // 2, size // 2
    r = size * 0.3
    br = size * 0.42

    for y in range(size):
        row = []
        for x in range(size):
            dx, dy = x - cx, y - cy
            dist = (dx*dx + dy*dy) ** 0.5
            # circle background
            if dist <= br:
                # simple bell shape: arc top + body
                in_bell = False
                # bell dome
                if dy < 0 and dist <= r * 0.85:
                    in_bell = True
                # bell body (rect below dome)
                if -r*0.1 <= dy <= r*0.55 and abs(dx) <= r*0.55:
                    in_bell = True
                # clapper
                if r*0.45 <= dy <= r*0.7 and abs(dx) <= r*0.18:
                    in_bell = True
                row.extend(fg if in_bell else bg)
            else:
                row.extend((255,255,255,0)[:3])
            row.append(255)  # alpha always 255
        pixels.append(bytes(row))

    raw = b''.join(b'\x00' + r for r in pixels)
    compressed = zlib.compress(raw, 9)

    ihdr = struct.pack('>IIBBBBB', size, size, 8, 6, 0, 0, 0)
    data  = chunk(b'IHDR', ihdr)
    data += chunk(b'IDAT', compressed)
    data += chunk(b'IEND', b'')
    return b'\x89PNG\r\n\x1a\n' + data

os.makedirs('public/icons', exist_ok=True)

for size in (16, 32, 64, 80):
    with open(f'public/icons/icon-{size}.png', 'wb') as f:
        f.write(png(size))
    print(f'  created icon-{size}.png')

print('Icons ready.')
