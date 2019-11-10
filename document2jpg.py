import os
import shutil
import subprocess

from pdf2image import convert_from_path as convert_to_jpg


def create_temp_pdf_dir(destination):
    import uuid
    temp_pdf_dir = os.path.join(destination, str(uuid.uuid4().hex))
    try:
        os.mkdir(temp_pdf_dir)
        return temp_pdf_dir
    except Exception:
        print('cannot create temp dir')
        exit(1)


def add_suffix(name):
    import datetime
    suffix = datetime.datetime.now().strftime("%d%m%y_%H%M%S")
    return "_".join([name, suffix])


if __name__ == "__main__":
    from sys import argv, exit

    if len(argv) < 3:
        print("USAGE: python %s <input-folder> <output-folder>" % argv[0])
        # exit(1)

    # input_folder = os.path.abspath(argv[1])
    # output_folder = os.path.abspath(argv[2])
    input_folder = os.path.abspath('demofiles')
    output_folder = os.path.abspath('outputfiles')

    if not os.path.isdir(input_folder):
        print("no such input folder: %s" % input_folder)
        exit(1)

    if not os.path.isdir(output_folder):
        try:
            os.mkdir(output_folder)
        except FileExistsError:
            print("ouput name: %s must be FOLDER!!!" % output_folder)
            exit(1)
    len_abspath = len(input_folder) + 1  # include closing slash

    queue = {}

    for root, dirs, files in os.walk(input_folder):
        subpath = root[len_abspath:]
        for name in files:
            filename, ext = os.path.splitext(name)
            filekey = os.path.join(subpath, filename)
            if filekey in queue:
                print(filename, 'exists in queue, trying to add suffix')
                filename = add_suffix(filename)
                filekey = os.path.join(subpath, filename)
                print('New name will be %s' % filename)

            queue[filekey] = {'subpath': subpath,
                              'filename': filename,
                              'ext': ext,
                              'orig_filename': name}

    temp_pdf_dir = create_temp_pdf_dir(output_folder)
    for key, value in queue.items():
        subpath = value['subpath']
        filename = value['filename']
        abspath_input = os.path.join(input_folder, subpath,
                                     value['orig_filename'])
        if value['ext'] == '.pdf':
            value['tmp_file'] = abspath_input
        elif value['ext'] in ('.jpg', '.jpeg'):
            value['tmp_file'] = None
        else:
            program = ['soffice', '--headless', '--convert-to', 'pdf']
            tmp_path = os.path.join(temp_pdf_dir, subpath, filename)
            process = subprocess.Popen(
                program + [abspath_input, '--outdir', tmp_path],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL)

            if process.wait() == 0:
                orig_filename, _ = os.path.splitext(value['orig_filename'])
                value['tmp_file'] = os.path.join(tmp_path, orig_filename + '.pdf')
            else:
                print('ERROR!', 'cant produce temp for %s' % abspath_input)

    for key, value in queue.items():
        subpath = value['subpath']
        try:
            tmp_file = value['tmp_file']
        except KeyError:
            print('ERROR!', 'cant convert %s in folder %s' %
                  (value['orig_filename'], subpath))
            continue
        filename = value['filename']

        jpg_dir = os.path.join(output_folder, subpath)
        if not os.path.exists(jpg_dir):
            os.makedirs(jpg_dir)
        jpg_name = os.path.join(jpg_dir, filename)

        print('Handle %s' % key)
        if tmp_file:
            pages = convert_to_jpg(tmp_file, 300)
            if len(pages) > 1:
                for page in pages:
                    print('== Save %s-page%d.jpg' % (jpg_name, pages.index(page)))
                    page.save('%s-page%d.jpg' %
                              (jpg_name, pages.index(page) + 1), 'JPEG')
            elif len(pages) == 1:
                print('== Save %s.jpg' % jpg_name)
                pages[0].save('%s.jpg' % jpg_name, 'JPEG')
        else:
            print('== Copy', '%s.jpg' % jpg_name)
            shutil.copy(os.path.join(input_folder, subpath, value['orig_filename']),
                        '%s.jpg' % jpg_name)

    # clean tmp_dir
    shutil.rmtree(temp_pdf_dir)
