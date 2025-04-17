import os
import re
from pathlib import Path
from tkinter import Tk, Button, filedialog, Label, messagebox
from tkinter.ttk import Progressbar
from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

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
            # For level 2 and above, keep the original folder name (e.g., "1.张娜拉")
            return folder_name
    return folder_name

def collect_files(folder_path, level=1):
    """Recursively collect folder structure and files for any level."""
    structure = []
    for item in sorted(Path(folder_path).iterdir(), key=lambda x: natural_sort_key(x.name)):
        if item.is_dir():
            # Check for subfolders (deeper levels)
            subfolders = collect_files(item, level + 1)
            # Check for image files in the current folder
            files = []
            for file in sorted(item.iterdir(), key=lambda x: natural_sort_key(x.name)):
                if file.suffix.lower() in ['.png', '.jpg', '.jpeg']:
                    files.append(file)
            # If there are subfolders or files, add to structure
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
    
    # Count total images for progress bar
    total_images = count_total_images(structure)
    current_image = 0
    
    # Content
    def add_headings(folder_structure, level=1, is_first_secondary=False):
        nonlocal current_image
        for i, (folder_name, files, subfolders) in enumerate(folder_structure, 1):
            # Convert folder name (e.g., "1.美女" to "一、美女", "1.张娜拉" kept as is)
            display_name = convert_folder_name(folder_name, level)
            
            # Add page break before secondary headings (except the first one after a primary heading)
            if level > 1 and not is_first_secondary:
                doc.add_page_break()
            
            # Add heading using paragraph to avoid automatic numbering
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(display_name)
            run.font.name = 'SimSun'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
            run.font.size = Pt(10.5)
            run.bold = (level == 1)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.style = f'Heading {level}'
            
            # Add images if any
            for file in files:
                print(f"Adding image: {file}")
                update_status(f"Adding image: {file}")
                try:
                    paragraph = doc.add_paragraph()
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = paragraph.add_run()
                    run.add_picture(str(file), width=Inches(6))
                    current_image += 1
                    update_progress(current_image / total_images * 100)
                except Exception as e:
                    print(f"Error adding image {file}: {e}")
            
            # Recursively add subfolders
            if subfolders:
                add_headings(subfolders, level + 1, is_first_secondary=(level == 1))

    add_headings(structure)
    
    # Save the document with overwrite check
    output_path_str = str(output_path)
    try:
        if os.path.exists(output_path_str):
            with open(output_path_str, 'a') as f:
                pass  # Check if file is writable
            doc.save(output_path_str)  # Overwrite existing file
        else:
            doc.save(output_path_str)
        print(f"Document creation completed at: {output_path_str}")
        update_status("文件创建完成！")
    except IOError:
        messagebox.showwarning("Warning", "请关闭文件再导出！")
        print("Failed to save: File is open or permission denied")
        update_status("Failed to save: File is open or permission denied")

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
            structure = collect_files(selected_folder)
            if not structure:
                messagebox.showerror("Error", "No valid files or folders found! Please check the folder structure.")
                return
            output_path = Path(selected_folder) / "佐证材料.docx"
            progress_bar['value'] = 0
            status_label.config(text="Starting document generation...")
            root.update()
            create_document(structure, selected_folder, output_path, 
                           lambda value: progress_bar.__setitem__('value', value) or root.update(),
                           lambda text: status_label.config(text=text) or root.update())
            label.config(text=f"文件生成位置: {output_path}")
            messagebox.showinfo("Success", f"Document generated at: {output_path}")
        except Exception as e:
            print(f"Unexpected error: {e}")
            if "Permission denied" not in str(e):
                messagebox.showerror("Error", f"Failed to generate document: {str(e)}")
            status_label.config(text=f"Error: {str(e)}")

    root = Tk()
    root.title("佐证材料生成")
    root.geometry("400x250")  # Increased height to accommodate progress bar
    
    label = Label(root, text="没有选择文件夹!")
    label.pack(pady=20)
    
    btn_select = Button(root, text="加载文件夹", command=select_folder)
    btn_select.pack(pady=10)
    
    btn_generate = Button(root, text="生成文件", command=generate_doc)
    btn_generate.pack(pady=10)
    
    # Add progress bar
    progress_bar = Progressbar(root, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=5)
    
    # Add status label
    status_label = Label(root, text="", wraplength=350)
    status_label.pack(pady=5)
    
    root.mainloop()

if __name__ == "__main__":
    main()