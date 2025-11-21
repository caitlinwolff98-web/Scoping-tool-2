#!/usr/bin/env python3
import os
import olefile
import struct

def decompress_stream(compressed):
    """Decompress VBA stream using RLE decompression"""
    decompressed = bytearray()
    pos = 0

    while pos < len(compressed):
        # Check if we have a flag byte
        if pos >= len(compressed):
            break

        flag = compressed[pos]
        pos += 1

        # Process based on flag
        for bit_index in range(8):
            if pos >= len(compressed):
                break

            if flag & (1 << bit_index):
                # Compressed token
                if pos + 1 >= len(compressed):
                    break
                token = struct.unpack('<H', compressed[pos:pos+2])[0]
                pos += 2

                # Extract offset and length
                offset = (token & 0xFFF) + 1
                length = (token >> 12) + 3

                # Copy from decompressed buffer
                for _ in range(length):
                    if len(decompressed) >= offset:
                        decompressed.append(decompressed[-offset])
                    else:
                        decompressed.append(0)
            else:
                # Literal byte
                if pos < len(compressed):
                    decompressed.append(compressed[pos])
                    pos += 1

    return bytes(decompressed)

def extract_vba_code(ole, stream_path):
    """Extract and decompress VBA code from a stream"""
    try:
        data = ole.openstream(stream_path).read()

        # VBA streams start with compression header
        if len(data) < 3:
            return None

        # Try to find the code section
        # VBA code is often stored after a header
        try:
            # Look for decompressed code patterns
            decoded = data.decode('latin-1', errors='ignore')
            if 'Sub ' in decoded or 'Function ' in decoded or 'Attribute' in decoded:
                return decoded
        except:
            pass

        # Try decompression
        try:
            # Skip compression header if present
            if data[0] == 1:  # Compression flag
                # Parse compressed data
                offset = 3
                decompressed = decompress_stream(data[offset:])
                decoded = decompressed.decode('latin-1', errors='ignore')
                if decoded:
                    return decoded
        except:
            pass

        # Try direct decode
        try:
            return data.decode('utf-16-le', errors='ignore')
        except:
            return data.decode('latin-1', errors='ignore')

    except Exception as e:
        print(f"Error extracting {stream_path}: {e}")
        return None

# Open the vbaProject.bin
vba_path = "vba_extracted/xl/vbaProject.bin"
ole = olefile.OleFileIO(vba_path)

# Extract each module
modules = {
    'ModConfig': 'VBA/__SRP_0',
    'ModDataProcessing': 'VBA/__SRP_1',
    'ModInteractiveDashboard': 'VBA/__SRP_2',
    'ModMain': 'VBA/__SRP_3',
    'ModManualScoping': 'VBA/__SRP_4',
    'ModPowerBIIntegration': 'VBA/__SRP_5',
    'ModSegmentAnalysis': 'VBA/__SRP_8',
    'ModTabCategorization': 'VBA/__SRP_9',
    'ModTableGeneration': 'VBA/__SRP_a',
    'ModThresholdScoping': 'VBA/__SRP_b',
    'Sheet1': 'VBA/__SRP_e',
    'ThisWorkbook': 'VBA/__SRP_f',
}

os.makedirs('vba_decoded', exist_ok=True)

for name, stream in modules.items():
    print(f"Extracting {name}...")
    code = extract_vba_code(ole, stream.split('/'))
    if code:
        output_file = f'vba_decoded/{name}.vba'
        with open(output_file, 'w', encoding='utf-8', errors='ignore') as f:
            f.write(code)
        print(f"  Saved to {output_file}")

        # Also try to get from the named stream
        try:
            named_code = extract_vba_code(ole, ['VBA', name])
            if named_code and len(named_code) > len(code):
                with open(output_file, 'w', encoding='utf-8', errors='ignore') as f:
                    f.write(named_code)
                print(f"  Updated from named stream")
        except:
            pass

ole.close()
print("\nVBA code extraction complete!")
