#!/usr/bin/env python3
import zipfile
import os
import olefile

# Extract vbaProject.bin from xlsm
xlsm_path = "Bidvest - Scoping Tool V12 (Links to PowerBI dashboard) (2).xlsm"
output_dir = "vba_extracted"

os.makedirs(output_dir, exist_ok=True)

# Extract vbaProject.bin
with zipfile.ZipFile(xlsm_path, 'r') as zip_ref:
    try:
        zip_ref.extract('xl/vbaProject.bin', output_dir)
        print(f"Extracted vbaProject.bin")
    except KeyError:
        print("No vbaProject.bin found")
        exit(1)

# Now parse the vbaProject.bin using olefile
vba_path = os.path.join(output_dir, 'xl', 'vbaProject.bin')

try:
    ole = olefile.OleFileIO(vba_path)

    # List all streams
    print("\nAvailable streams:")
    for stream in ole.listdir():
        print(f"  {'/'.join(stream)}")

    # Try to extract VBA modules
    vba_dir = output_dir + "/modules"
    os.makedirs(vba_dir, exist_ok=True)

    # Common VBA module paths
    vba_paths = [
        'VBA',
        '_VBA_PROJECT_CUR',
        'PROJECT',
        'dir'
    ]

    # Extract all streams that might contain VBA code
    for stream in ole.listdir():
        stream_path = '/'.join(stream)
        if 'VBA' in stream_path or any(x in stream_path for x in ['Module', 'Class', 'Form', 'Sheet', 'ThisWorkbook']):
            try:
                data = ole.openstream(stream).read()
                filename = stream[-1].replace('/', '_')
                output_path = os.path.join(vba_dir, filename)
                with open(output_path, 'wb') as f:
                    f.write(data)
                print(f"Extracted: {stream_path} -> {filename}")
            except Exception as e:
                print(f"Could not extract {stream_path}: {e}")

    ole.close()
    print(f"\nVBA modules extracted to {vba_dir}")

except Exception as e:
    print(f"Error parsing vbaProject.bin: {e}")
    exit(1)
