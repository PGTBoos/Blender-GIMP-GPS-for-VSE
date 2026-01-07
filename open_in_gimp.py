"""
VSE Tools - Blender VSE Addon
Tools for working with image strips in the Video Sequence Editor:
- Open in GIMP for external editing
- Open source folder in file browser  
- Show photo location in Google Maps (if GPS data exists)

No external dependencies - uses pure Python to read EXIF GPS data.

Installation:
1. Edit → Preferences → Add-ons
2. Click dropdown arrow → "Install from Disk..."
3. Select this .py file
4. Enable the addon "Sequencer: VSE Tools"

Usage:
- Select an image strip in the VSE
- Right-click to see options, or use keyboard shortcuts:
  - Ctrl+Shift+G: Open in GIMP
  - Ctrl+Shift+F: Open source folder
  - Ctrl+Shift+M: Show in Google Maps (if GPS available)
"""

bl_info = {
    "name": "Sequencer: VSE Tools",
    "author": "Peter",
    "version": (1, 2, 0),
    "blender": (5, 0, 0),
    "location": "Sequencer > Strip menu, or Right-click menu",
    "description": "Open images in GIMP, open source folder, show GPS location in Google Maps",
    "category": "Sequencer",
}

import bpy
import subprocess
import os
import platform
import webbrowser
import struct


# =============================================================================
# PURE PYTHON EXIF GPS EXTRACTION (No external dependencies)
# =============================================================================

def read_jpeg_exif_gps(filepath):
    """
    Read GPS coordinates from JPEG EXIF data using pure Python.
    Returns (latitude, longitude) as floats, or None if not found.
    """
    try:
        with open(filepath, 'rb') as f:
            # Check JPEG magic bytes
            if f.read(2) != b'\xff\xd8':
                return None
            
            # Find EXIF APP1 marker
            while True:
                marker = f.read(2)
                if len(marker) < 2:
                    return None
                
                if marker == b'\xff\xe1':  # APP1 (EXIF)
                    break
                elif marker[0:1] == b'\xff':
                    # Skip other markers
                    length = struct.unpack('>H', f.read(2))[0]
                    f.seek(length - 2, 1)
                else:
                    return None
            
            # Read APP1 length and EXIF header
            length = struct.unpack('>H', f.read(2))[0]
            exif_header = f.read(6)
            
            if exif_header != b'Exif\x00\x00':
                return None
            
            # TIFF header starts here
            tiff_start = f.tell()
            tiff_header = f.read(8)
            
            # Determine byte order
            if tiff_header[0:2] == b'II':
                endian = '<'  # Little endian
            elif tiff_header[0:2] == b'MM':
                endian = '>'  # Big endian
            else:
                return None
            
            # Get offset to first IFD
            ifd_offset = struct.unpack(endian + 'I', tiff_header[4:8])[0]
            
            # Helper function to read values
            def read_ifd_entries(offset):
                f.seek(tiff_start + offset)
                num_entries = struct.unpack(endian + 'H', f.read(2))[0]
                entries = {}
                
                for _ in range(num_entries):
                    tag = struct.unpack(endian + 'H', f.read(2))[0]
                    type_id = struct.unpack(endian + 'H', f.read(2))[0]
                    count = struct.unpack(endian + 'I', f.read(4))[0]
                    value_offset = f.read(4)
                    
                    entries[tag] = (type_id, count, value_offset)
                
                return entries
            
            def get_value(type_id, count, value_offset):
                """Get actual value from IFD entry."""
                type_sizes = {1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 7: 1, 9: 4, 10: 8}
                size = type_sizes.get(type_id, 1) * count
                
                if size <= 4:
                    data = value_offset
                else:
                    offset = struct.unpack(endian + 'I', value_offset)[0]
                    pos = f.tell()
                    f.seek(tiff_start + offset)
                    data = f.read(size)
                    f.seek(pos)
                
                return data
            
            def read_rational(data, offset=0):
                """Read a rational (fraction) value."""
                if isinstance(data, bytes):
                    num = struct.unpack(endian + 'I', data[offset:offset+4])[0]
                    den = struct.unpack(endian + 'I', data[offset+4:offset+8])[0]
                else:
                    return 0
                return num / den if den != 0 else 0
            
            def read_gps_coord(data):
                """Read GPS coordinate (3 rationals: degrees, minutes, seconds)."""
                degrees = read_rational(data, 0)
                minutes = read_rational(data, 8)
                seconds = read_rational(data, 16)
                return degrees + minutes / 60 + seconds / 3600
            
            # Read IFD0
            ifd0 = read_ifd_entries(ifd_offset)
            
            # Find GPS IFD pointer (tag 0x8825)
            if 0x8825 not in ifd0:
                return None
            
            gps_offset = struct.unpack(endian + 'I', ifd0[0x8825][2])[0]
            
            # Read GPS IFD
            gps_ifd = read_ifd_entries(gps_offset)
            
            # GPS tags we need:
            # 0x0001 = GPSLatitudeRef (N/S)
            # 0x0002 = GPSLatitude
            # 0x0003 = GPSLongitudeRef (E/W)
            # 0x0004 = GPSLongitude
            
            if 0x0002 not in gps_ifd or 0x0004 not in gps_ifd:
                return None
            
            # Read latitude
            lat_data = get_value(*gps_ifd[0x0002])
            latitude = read_gps_coord(lat_data)
            
            # Read latitude reference
            if 0x0001 in gps_ifd:
                lat_ref = get_value(*gps_ifd[0x0001])
                if isinstance(lat_ref, bytes) and lat_ref[0:1] in (b'S', b's'):
                    latitude = -latitude
            
            # Read longitude
            lon_data = get_value(*gps_ifd[0x0004])
            longitude = read_gps_coord(lon_data)
            
            # Read longitude reference
            if 0x0003 in gps_ifd:
                lon_ref = get_value(*gps_ifd[0x0003])
                if isinstance(lon_ref, bytes) and lon_ref[0:1] in (b'W', b'w'):
                    longitude = -longitude
            
            return (latitude, longitude)
    
    except Exception as e:
        print(f"Error reading EXIF GPS: {e}")
        return None


def get_gps_coordinates(filepath):
    """Get GPS coordinates from an image file. Returns (lat, lon) or None."""
    # Check file extension
    ext = os.path.splitext(filepath)[1].lower()
    
    if ext in ('.jpg', '.jpeg'):
        return read_jpeg_exif_gps(filepath)
    
    # For other formats, we could add support later
    # PNG doesn't typically have EXIF, HEIC needs different parsing
    return None


# =============================================================================
# GIMP PATH DETECTION
# =============================================================================

def get_default_gimp_path():
    """Try to find GIMP installation path based on OS."""
    system = platform.system()
    
    if system == "Windows":
        possible_paths = [
            r"C:\Program Files\GIMP 2\bin\gimp-2.10.exe",
            r"C:\Program Files\GIMP 2\bin\gimp-2.99.exe",
            r"C:\Program Files\GIMP 3\bin\gimp-3.0.exe",
            r"C:\Program Files (x86)\GIMP 2\bin\gimp-2.10.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\GIMP 2\bin\gimp-2.10.exe"),
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return "gimp"
    
    elif system == "Darwin":  # macOS
        possible_paths = [
            "/Applications/GIMP.app/Contents/MacOS/gimp",
            "/Applications/GIMP-2.10.app/Contents/MacOS/gimp",
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return "gimp"
    
    else:  # Linux
        return "gimp"


# =============================================================================
# HELPER FUNCTION TO GET FILE PATH FROM STRIP
# =============================================================================

def get_strip_filepath(strip, context):
    """Get the file path from an image or movie strip."""
    if strip.type == 'IMAGE':
        directory = bpy.path.abspath(strip.directory)
        if strip.elements:
            # For single-image strips, just get the first element
            if len(strip.elements) == 1:
                filename = strip.elements[0].filename
            else:
                # For multi-image strips, get current frame's image
                current_frame = context.scene.frame_current
                strip_start = strip.frame_final_start
                element_index = current_frame - strip_start
                element_index = max(0, min(element_index, len(strip.elements) - 1))
                filename = strip.elements[element_index].filename
            return os.path.join(directory, filename)
    
    elif strip.type == 'MOVIE':
        return bpy.path.abspath(strip.filepath)
    
    return None


# =============================================================================
# ADDON PREFERENCES
# =============================================================================

class VSEToolsPreferences(bpy.types.AddonPreferences):
    bl_idname = __package__ if __package__ else __name__
    
    gimp_path: bpy.props.StringProperty(
        name="GIMP Executable Path",
        description="Path to GIMP executable. Leave empty for auto-detect",
        default="",
        subtype='FILE_PATH'
    )
    
    def draw(self, context):
        layout = self.layout
        layout.prop(self, "gimp_path")
        layout.label(text="Leave empty to auto-detect GIMP location")
        
        if not self.gimp_path:
            detected = get_default_gimp_path()
            layout.label(text=f"Auto-detected: {detected}", icon='INFO')


# =============================================================================
# OPERATORS
# =============================================================================

class SEQUENCER_OT_open_in_gimp(bpy.types.Operator):
    """Open the selected strip's source image in GIMP"""
    bl_idname = "sequencer.open_in_gimp"
    bl_label = "Open in GIMP"
    bl_options = {'REGISTER', 'UNDO'}
    
    @classmethod
    def poll(cls, context):
        if context.scene.sequence_editor is None:
            return False
        strip = context.scene.sequence_editor.active_strip
        if strip is None:
            return False
        return strip.type in {'IMAGE', 'MOVIE'}
    
    def execute(self, context):
        strip = context.scene.sequence_editor.active_strip
        filepath = get_strip_filepath(strip, context)
        
        if not filepath:
            self.report({'ERROR'}, "Could not determine file path")
            return {'CANCELLED'}
        
        if not os.path.exists(filepath):
            self.report({'ERROR'}, f"File not found: {filepath}")
            return {'CANCELLED'}
        
        # Get GIMP path
        addon_name = __package__ if __package__ else __name__
        prefs = context.preferences.addons[addon_name].preferences
        gimp_path = prefs.gimp_path if prefs.gimp_path else get_default_gimp_path()
        
        try:
            subprocess.Popen([gimp_path, filepath])
            self.report({'INFO'}, f"Opened in GIMP: {os.path.basename(filepath)}")
        except FileNotFoundError:
            self.report({'ERROR'}, f"GIMP not found at: {gimp_path}\nSet the correct path in addon preferences")
            return {'CANCELLED'}
        except Exception as e:
            self.report({'ERROR'}, f"Failed to open GIMP: {str(e)}")
            return {'CANCELLED'}
        
        return {'FINISHED'}


class SEQUENCER_OT_open_source_folder(bpy.types.Operator):
    """Open the folder containing the strip's source file"""
    bl_idname = "sequencer.open_source_folder"
    bl_label = "Open Source Folder"
    bl_options = {'REGISTER'}
    
    @classmethod
    def poll(cls, context):
        if context.scene.sequence_editor is None:
            return False
        strip = context.scene.sequence_editor.active_strip
        if strip is None:
            return False
        return strip.type in {'IMAGE', 'MOVIE', 'SOUND'}
    
    def execute(self, context):
        strip = context.scene.sequence_editor.active_strip
        
        if strip.type == 'IMAGE':
            directory = bpy.path.abspath(strip.directory)
        elif strip.type in {'MOVIE', 'SOUND'}:
            directory = os.path.dirname(bpy.path.abspath(strip.filepath))
        else:
            self.report({'ERROR'}, "Unsupported strip type")
            return {'CANCELLED'}
        
        if not os.path.exists(directory):
            self.report({'ERROR'}, f"Folder not found: {directory}")
            return {'CANCELLED'}
        
        system = platform.system()
        try:
            if system == "Windows":
                os.startfile(directory)
            elif system == "Darwin":
                subprocess.Popen(["open", directory])
            else:
                subprocess.Popen(["xdg-open", directory])
            
            self.report({'INFO'}, f"Opened folder: {directory}")
        except Exception as e:
            self.report({'ERROR'}, f"Failed to open folder: {str(e)}")
            return {'CANCELLED'}
        
        return {'FINISHED'}


class SEQUENCER_OT_show_in_google_maps(bpy.types.Operator):
    """Open the photo's GPS location in Google Maps (if GPS data exists in JPEG)"""
    bl_idname = "sequencer.show_in_google_maps"
    bl_label = "Show in Google Maps"
    bl_options = {'REGISTER'}
    
    @classmethod
    def poll(cls, context):
        if context.scene.sequence_editor is None:
            return False
        strip = context.scene.sequence_editor.active_strip
        if strip is None:
            return False
        return strip.type == 'IMAGE'
    
    def execute(self, context):
        strip = context.scene.sequence_editor.active_strip
        filepath = get_strip_filepath(strip, context)
        
        if not filepath:
            self.report({'ERROR'}, "Could not determine file path")
            return {'CANCELLED'}
        
        if not os.path.exists(filepath):
            self.report({'ERROR'}, f"File not found: {filepath}")
            return {'CANCELLED'}
        
        # Check if it's a JPEG
        ext = os.path.splitext(filepath)[1].lower()
        if ext not in ('.jpg', '.jpeg'):
            self.report({'WARNING'}, f"GPS reading only supported for JPEG files (this is {ext})")
            return {'CANCELLED'}
        
        # Get GPS coordinates
        coords = get_gps_coordinates(filepath)
        
        if coords is None:
            self.report({'WARNING'}, f"No GPS data found in: {os.path.basename(filepath)}")
            return {'CANCELLED'}
        
        lat, lon = coords
        
        # Open Google Maps
        url = f"https://www.google.com/maps?q={lat},{lon}"
        webbrowser.open(url)
        
        self.report({'INFO'}, f"GPS: {lat:.6f}, {lon:.6f}")
        
        return {'FINISHED'}


# =============================================================================
# MENUS
# =============================================================================

def draw_strip_menu(self, context):
    """Add to Strip menu"""
    layout = self.layout
    layout.separator()
    layout.operator("sequencer.open_in_gimp", icon='IMAGE_DATA')
    layout.operator("sequencer.open_source_folder", icon='FILE_FOLDER')
    layout.operator("sequencer.show_in_google_maps", icon='URL')


def draw_context_menu(self, context):
    """Add to right-click context menu"""
    layout = self.layout
    layout.separator()
    layout.operator("sequencer.open_in_gimp", icon='IMAGE_DATA')
    layout.operator("sequencer.open_source_folder", icon='FILE_FOLDER')
    layout.operator("sequencer.show_in_google_maps", icon='URL')


# =============================================================================
# KEYMAPS
# =============================================================================

addon_keymaps = []


def register_keymaps():
    wm = bpy.context.window_manager
    kc = wm.keyconfigs.addon
    
    if kc:
        km = kc.keymaps.new(name='Sequencer', space_type='SEQUENCE_EDITOR')
        
        # Ctrl+Shift+G to open in GIMP
        kmi = km.keymap_items.new(
            "sequencer.open_in_gimp",
            type='G',
            value='PRESS',
            ctrl=True,
            shift=True
        )
        addon_keymaps.append((km, kmi))
        
        # Ctrl+Shift+F to open folder
        kmi = km.keymap_items.new(
            "sequencer.open_source_folder",
            type='F',
            value='PRESS',
            ctrl=True,
            shift=True
        )
        addon_keymaps.append((km, kmi))
        
        # Ctrl+Shift+M to show in Google Maps
        kmi = km.keymap_items.new(
            "sequencer.show_in_google_maps",
            type='M',
            value='PRESS',
            ctrl=True,
            shift=True
        )
        addon_keymaps.append((km, kmi))


def unregister_keymaps():
    for km, kmi in addon_keymaps:
        km.keymap_items.remove(kmi)
    addon_keymaps.clear()


# =============================================================================
# REGISTRATION
# =============================================================================

classes = [
    VSEToolsPreferences,
    SEQUENCER_OT_open_in_gimp,
    SEQUENCER_OT_open_source_folder,
    SEQUENCER_OT_show_in_google_maps,
]


def register():
    for cls in classes:
        bpy.utils.register_class(cls)
    
    bpy.types.SEQUENCER_MT_strip.append(draw_strip_menu)
    bpy.types.SEQUENCER_MT_context_menu.append(draw_context_menu)
    
    register_keymaps()


def unregister():
    unregister_keymaps()
    
    bpy.types.SEQUENCER_MT_strip.remove(draw_strip_menu)
    bpy.types.SEQUENCER_MT_context_menu.remove(draw_context_menu)
    
    for cls in reversed(classes):
        bpy.utils.unregister_class(cls)


if __name__ == "__main__":
    register()
