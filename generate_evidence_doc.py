import os
import re
from pathlib import Path
from tkinter import Tk, Button, filedialog, Label, messagebox
from tkinter.ttk import Progressbar
from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from PIL import Image

# Global variable to store the selected folder
selected_folder = ""

def natural_sort_key(s):
    """Extract numbers from filename for natural sorting."""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', s)]

def to_chinese_num(num):
    """Convert number to Chinese numeral with a comma."""
    chinese_nums = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
    if num <= 10:
        return f"{chinese_nums[num]}、"
    elif num < 20:
        return f"十{chinese_nums[num - 10]}、"
    else:
        return f"{num}、"

def convert_folder_name(folder_name, level):
    """Convert folder name like '1.美女' to '一、美女' for level 1, keep original for other levels."""
    match = re.match(r'(\d+)\.(.*)', folder_name)
    if match:
        num = int(match.group(1))
        rest = match.group(2).strip()
        if level == 1:
            return f"{to_chinese_num(num)}{rest}"
        else:
            return folder_name
    return folder_name

def rename_images(folder_path, update_status):
    """Recursively rename images in all folders to 1.jpg, 2.jpg, etc., preserving extension."""
    for item in Path(folder_path).iterdir():
        if item.is_dir():
            rename_images(item, update_status)
        elif item.is_file() and item.suffix.lower() in ['.png', '.jpg', '.jpeg']:
            images = [f for f in item.parent.iterdir() if f.suffix.lower() in ['.png', '.jpg', '.jpeg']]
            images = sorted(images, key=lambda x: natural_sort_key(x.name))
            for i, image in enumerate(images, 1):
                new_name = item.parent / f"{i}{image.suffix.lower()}"
                if image != new_name:
                    if new_name.exists():
                        update_status(f"Warning: {new_name} already exists, skipping rename for {image}")
                        print(f"Warning: {new_name} already exists, skipping rename for {image}")
                        continue
                    try:
                        image.rename(new_name)
                        update_status(f"Renamed {image} to {new_name}")
                        print(f"Renamed {image} to {new_name}")
                    except Exception as e:
                        update_status(f"Error renaming {image}: {e}")
                        print(f"Error renaming {image}: {e}")

def collect_files(folder_path, level=1):
    """Recursively collect folder structure and files for any level."""
    structure = []
    for item in sorted(Path(folder_path).iterdir(), key=lambda x: natural_sort_key(x.name)):
        if item.is_dir():
            subfolders = collect_files(item, level + 1)
            files = []
            for file in sorted(item.iterdir(), key=lambda x: natural_sort_key(x.name)):
                if file.suffix.lower() in ['.png', '.jpg', '.jpeg']:
                    files.append(file)
            if subfolders or files:
                structure.append((item.name, files, subfolders))
    return structure

def count_total_images(structure):
    """Count total images in the structure for progress bar."""
    total = 0
    for folder_name, files, subfolders in structure:
        total += len(files)
        if subfolders:
            total += count_total_images(subfolders)
    return total

def get_image_dimensions(image_path):
    """Get the dimensions of an image using PIL."""
    try:
        with Image.open(image_path) as img:
            width, height = img.size
            return width, height
    except Exception as e:
        print(f"Error getting image dimensions for {image_path}: {e}")
        return None, None

def create_document(structure, root_path, output_path, update_progress, update_status):
    """Create Word document with content pages, supporting multiple levels."""
    print(f"Creating document at: {output_path}")
    doc = Document()
    
    # Set page margins for all pages
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(1.27)
        section.left_margin = Cm(1.27)
        section.bottom_margin = Cm(1.27)
        section.right_margin = Cm(1.27)
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
    
    # Calculate available dimensions
    page_width = 21  # cm (A4 width)
    page_height = 29.7  # cm (A4 height)
    margin = 1.27  # cm
    available_width = page_width - 2 * margin  # cm
    available_width_in = available_width / 2.54  # inches (18.46 cm / 2.54 ≈ 7.27 inches)
    available_height = page_height - 2 * margin  # cm
    available_height_in = available_height / 2.54  # inches (27.16 cm / 2.54 ≈ 10.69 inches)
    
    # Estimate title height (assume 1 line per title, font size 10.5 pt, 1 pt = 1/72 inch)
    title_height_in = (10.5 / 72) * 2  # Increased to 2 lines for safety
    
    # Scaling factor to leave some margin (95% of calculated size)
    scale_factor = 0.95
    
    # Customize Heading 1 and Heading 2 styles
    styles = doc.styles
    style1 = styles['Heading 1']
    style1.font.name = 'SimSun'
    style1._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
    style1.font.size = Pt(10.5)
    style1.font.bold = True
    style1.paragraph_format.space_before = Pt(0)
    style1.paragraph_format.space_after = Pt(0)
    style1.paragraph_format.line_spacing = 1.0
    style1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    style2 = styles['Heading 2']
    style2.font.name = 'SimSun'
    style2._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
    style2.font.size = Pt(10.5)
    style2.font.bold = False
    style2.paragraph_format.space_before = Pt(0)
    style2.paragraph_format.space_after = Pt(0)
    style2.paragraph_format.line_spacing = 1.0
    style2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    total_images = count_total_images(structure)
    current_image = 0
    
    def add_headings(folder_structure, level=1, is_first_secondary=False):
        nonlocal current_image
        for i, (folder_name, files, subfolders) in enumerate(folder_structure, 1):
            display_name = convert_folder_name(folder_name, level)
            
            # Add heading
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(display_name)
            run.font.name = 'SimSun'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
            run.font.size = Pt(10.5)
            run.bold = (level == 1)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.style = f'Heading {level}'
            # Ensure heading stays with the next content
            paragraph.paragraph_format.keep_with_next = True
            
            # Add images if any
            for file in files:
                print(f"Adding image: {file}")
                update_status(f"Adding image: {file}")
                try:
                    # Get image dimensions
                    img_width, img_height = get_image_dimensions(file)
                    if img_width is None or img_height is None:
                        image_width = available_width_in * scale_factor
                        image_height = None
                    else:
                        # Calculate aspect ratio
                        aspect_ratio = img_width / img_height
                        
                        # Step 1: Try width-first (fit to page width)
                        image_width = available_width_in
                        image_height = image_width / aspect_ratio
                        
                        # Step 2: Check if height fits (including title height)
                        total_height = image_height + title_height_in
                        if total_height > available_height_in:
                            # Step 3: Height-first (fit to page height)
                            image_height = available_height_in - title_height_in
                            image_width = image_height * aspect_ratio
                            # Ensure width doesn't exceed page width
                            if image_width > available_width_in:
                                image_width = available_width_in
                                image_height = image_width / aspect_ratio
                        
                        # Step 4: Apply scaling factor to leave margin
                        image_width *= scale_factor
                        image_height *= scale_factor
                    
                    # Add image with calculated dimensions
                    paragraph = doc.add_paragraph()
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = paragraph.add_run()
                    if image_height:
                        run.add_picture(str(file), width=Inches(image_width), height=Inches(image_height))
                    else:
                        run.add_picture(str(file), width=Inches(image_width))
                    current_image += 1
                    update_progress(current_image / total_images * 100)
                except Exception as e:
                    print(f"Error adding image {file}: {e}")
            
            # Recursively add subfolders
            if subfolders:
                if level > 1 and not is_first_secondary:
                    doc.add_page_break()
                add_headings(subfolders, level + 1, is_first_secondary=(level == 1))

    add_headings(structure)
    
    output_path_str = str(output_path)
    try:
        if os.path.exists(output_path_str):
            with open(output_path_str, 'a') as f:
                pass
        doc.save(output_path_str)
        print(f"Document creation completed at: {output_path_str}")
        update_status("文件创建完成！")
        return True
    except IOError:
        messagebox.showwarning("Warning", "请关闭文件再导出！")
        print("Failed to save: File is open or permission denied")
        update_status("Failed to save: File is open or permission denied")
        return False

def main():
    def select_folder():
        global selected_folder
        folder_path = filedialog.askdirectory()
        if folder_path:
            selected_folder = folder_path
            label.config(text=f"Selected: {folder_path}")
            print(f"Folder selected: {selected_folder}")
        else:
            label.config(text="没有选择文件夹!")
            print("没有选择文件夹!")

    def generate_doc():
        if not selected_folder:
            label.config(text="请选择文件夹！")
            print("没有选择文件夹!")
            return
        try:
            if messagebox.askyesno("确认重命名", "是否要将所有图片重命名为 1.jpg, 2.jpg 等？此操作不可逆！"):
                status_label.config(text="正在重命名图片...")
                root.update()
                rename_images(selected_folder, lambda text: status_label.config(text=text) or root.update())
                status_label.config(text="图片重命名完成！")
            
            structure = collect_files(selected_folder)
            if not structure:
                messagebox.showerror("Error", "No valid files or folders found! Please check the folder structure.")
                return
            output_path = Path(selected_folder) / "佐证材料.docx"
            progress_bar['value'] = 0
            status_label.config(text="Starting document generation...")
            root.update()
            success = create_document(structure, selected_folder, output_path, 
                                     lambda value: progress_bar.__setitem__('value', value) or root.update(),
                                     lambda text: status_label.config(text=text) or root.update())
            if success:
                label.config(text=f"文件生成位置: {output_path}")
                messagebox.showinfo("Success", f"Document generated at: {output_path}")
        except Exception as e:
            print(f"Unexpected error: {e}")
            if "Permission denied" not in str(e):
                messagebox.showerror("Error", f"Failed to generate document: {str(e)}")
            status_label.config(text=f"Error: {str(e)}")

    root = Tk()
    root.title("佐证材料生成")
    root.geometry("400x250")
    
    label = Label(root, text="没有选择文件夹!")
    label.pack(pady=20)
    
    btn_select = Button(root, text="加载文件夹", command=select_folder)
    btn_select.pack(pady=10)
    
    btn_generate = Button(root, text="生成文件", command=generate_doc)
    btn_generate.pack(pady=10)
    
    progress_bar = Progressbar(root, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=5)
    
    status_label = Label(root, text="", wraplength=350)
    status_label.pack(pady=5)
    
    root.mainloop()

if __name__ == "__main__":
    main()