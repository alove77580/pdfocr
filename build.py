import os
import shutil
import subprocess

def copy_dlls(src_dir, dst_dir, dll_names):
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)
    
    for dll in dll_names:
        src = os.path.join(src_dir, dll)
        dst = os.path.join(dst_dir, dll)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"已复制: {dll}")
        else:
            print(f"警告: 找不到 {dll}")

def copy_directory(src_dir, dst_dir):
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)
    
    for item in os.listdir(src_dir):
        src = os.path.join(src_dir, item)
        dst = os.path.join(dst_dir, item)
        if os.path.isfile(src):
            shutil.copy2(src, dst)
            print(f"已复制: {item}")
        elif os.path.isdir(src):
            copy_directory(src, dst)

def build_exe():
    # 确保tesseract目录存在
    if not os.path.exists('tesseract'):
        os.makedirs('tesseract')
    
    # 复制Tesseract文件
    tesseract_install_dir = r'D:\Program Files\Tesseract-OCR'  # 默认安装路径
    if os.path.exists(tesseract_install_dir):
        # 复制tesseract.exe
        shutil.copy2(
            os.path.join(tesseract_install_dir, 'tesseract.exe'),
            os.path.join('tesseract', 'tesseract.exe')
        )
        
        # 复制tessdata目录
        tessdata_src = os.path.join(tesseract_install_dir, 'tessdata')
        tessdata_dst = os.path.join('tesseract', 'tessdata')
        if os.path.exists(tessdata_src):
            copy_directory(tessdata_src, tessdata_dst)
        
        # 复制Tesseract的DLL文件
        tesseract_dlls = [
            'libtesseract-5.dll',
            'libgcc_s_seh-1.dll',
            'libstdc++-6.dll',
            'libwinpthread-1.dll',
            'zlib1.dll',
            'libpng16-16.dll',
            'libjpeg-8.dll',
            'libtiff-6.dll',
            'libwebp-7.dll'
        ]
        copy_dlls(tesseract_install_dir, 'tesseract', tesseract_dlls)
    
    # 复制Poppler文件
    poppler_install_dir = r'D:\Program Files\poppler-24.08.0\Library\bin'  # 请根据实际安装路径修改
    if os.path.exists(poppler_install_dir):
        poppler_dst_dir = 'poppler'
        if not os.path.exists(poppler_dst_dir):
            os.makedirs(poppler_dst_dir)
        
        # 复制所有Poppler文件
        copy_directory(poppler_install_dir, poppler_dst_dir)
        
        # 复制Poppler的DLL文件
        poppler_dlls = [
            'libpoppler-126.dll',
            'libjpeg-8.dll',
            'libpng16-16.dll',
            'libtiff-6.dll',
            'libwebp-7.dll',
            'zlib1.dll',
            'libwinpthread-1.dll',
            'libgcc_s_seh-1.dll',
            'libstdc++-6.dll'
        ]
        copy_dlls(poppler_install_dir, 'poppler', poppler_dlls)
    
    # 使用PyInstaller打包
    subprocess.run([
        'pyinstaller',
        '--name=PDF_OCR_Tool',
        '--windowed',
        '--add-data=tesseract;tesseract',
        '--add-data=poppler;poppler',
        '--hidden-import=PIL._tkinter',
        '--hidden-import=pytesseract',
        '--hidden-import=pdf2image',
        '--hidden-import=PIL',
        '--hidden-import=PIL.Image',
        '--hidden-import=PIL.ImageTk',
        '--clean',  # 清理临时文件
        'main.py'
    ])

    # 清理临时目录
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('tesseract'):
        shutil.rmtree('tesseract')
    if os.path.exists('poppler'):
        shutil.rmtree('poppler')
    if os.path.exists('PDF_OCR_Tool.spec'):
        os.remove('PDF_OCR_Tool.spec')

if __name__ == '__main__':
    build_exe() 