import os
import subprocess
import sys

__author__ = 'Avery'

# is this system 32-bit or 64-bit
SYS_BITS = sys.maxsize.bit_length() + 1
if SYS_BITS not in {32, 64}:
    raise OSError(f'could not determine if system is 32-bit or 64-bit: {SYS_BITS}')

# first try to find the 64-bit executable
if SYS_BITS == 64:
    EXE_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), f'7-zip/x64/7z.exe'))
else:
    EXE_PATH = ''

# fallback to 32-bit
if not os.path.exists(EXE_PATH):
    EXE_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), f'7-zip/x32/7z.exe'))

# double-check
assert os.path.exists(EXE_PATH), 'could not find 7-zip executable'


def archive_create(files_and_folders, archive, password=None, encrypt_headers=False, overwrite=False, verbose=3,
                   volumes=None):
    """
    create 7z archive from some files and folders
    files and folders will be placed in the root of the archive

    :param files_and_folders: list of paths
    :param archive: path (or name)
    :param password: ascii only, excluding null bytes and double-quote char
    :param encrypt_headers: encrypt file names and directory tree within the archive
    :param overwrite: overwrite existing file (but not dir)
    :param verbose: loudness as int {0, 1, 2, 3}
    :param volumes: size of volume, eg '10k' or '2g'
    :return: output printed by 7zip
    """

    # make paths absolute and unicode
    archive_path = os.path.abspath(archive)
    todo_paths = [os.path.abspath(item) for item in files_and_folders]

    # check validity of input files and folders
    seen = dict()
    for item_path in files_and_folders:
        seen.setdefault(os.path.basename(item_path), set()).add(item_path)

    for item_name, item_paths in seen.items():
        if len(item_paths) > 1:
            print(f'multiple items with same filename added to archive: <{item_name}>', file=sys.stderr)
            for item_path in item_paths:
                print(f'<{item_path}>', file=sys.stderr)
            raise ValueError(item_name)

    # 7z will not create an archive over an existing dir
    assert not os.path.isdir(archive_path), 'dir already exists at output path'

    # don't want to accidentally tell 7zip to overwrite anything
    assert overwrite or not os.path.isfile(archive_path), 'file already exists at output path'

    # base command (without input files/folders)
    command = [EXE_PATH,
               'a',
               archive_path,
               '-t7z',
               '-m0=lzma2',
               '-mx=9',
               f'-bb{verbose}',
               '-bt',
               ]

    # -t7z         -- type of archive             -> 7z
    # -m0=lzma2    -- compression algorithm       -> lzma2
    # -mx=9        -- compression level           -> 9 = ultra
    # -aoa         -- duplicate destination files -> oa = overwrite all
    # -mfb=64      -- fast bytes                  -> what is this
    # -md=32m      -- dictionary size?            -> 32Mb
    # -ms=on       -- solid                       -> yes (default)
    # -d=64m       -- dictionary size             -> 66Mb
    # -mhe         -- encrypt header              -> default off
    # -p{PASSWORD} -- set password = "{PASSWORD}" -> default unencrypted

    # overwrite
    if overwrite:
        command += ['-aoa']

    # split into volumes (bytes, kilobytes, megabytes, gigabytes)
    # 7z a a.7z *.txt -v10k -v15k -v2m
    # First volume will be 10 KB, second will be 15 KB, and all others will be 2 MB.
    if volumes is not None:
        assert type(volumes) is str and len(volumes) > 0
        for size in volumes.strip().split():
            command += [f'-v{size}']

    # add password and header encryption
    if password is not None:
        assert len(password) > 0, 'password must be at least one character long'

        # need to write and run batch file to use this char
        if '"' in password:
            raise NotImplementedError('double-quote not supported')

        # command += u' "-p{PASSWORD}"'.format(PASSWORD=password)
        command += [f'-p{password}']
        if encrypt_headers:
            command += ['-mhe']

    # add files
    command += todo_paths

    # make the file, check that it's okay
    ret_val = subprocess.check_output(command).decode('cp1252')
    assert 'Everything is Ok' in ret_val, f'something went wrong: {ret_val}'

    # TODO: parse ret_val into something useful
    return ret_val


def archive_test(archive, password=None, verbose=3):
    """
    test the integrity of an archive

    :param archive: path (or name)
    :param password: ascii only, excluding null bytes and double-quote char
    :param verbose: loudness as int {0, 1, 2, 3}
    :return: output printed by 7zip
    """
    archive_path = os.path.abspath(archive)

    # set arbitrary password if none given
    if password is None:
        password = '\x7f'  # ascii for DEL key

    # validity of provided password (if any)
    assert len(password) > 0, 'password must be at least one character long'

    # need to write and run batch file to use this char
    if '"' in password:
        raise NotImplementedError('double-quote not supported')

    # always supply a password
    command = [EXE_PATH,
               't',
               archive_path,
               f'-p{password}',
               f'-bb{verbose}',
               '-bt',
               ]

    # make the file, check that it's okay
    ret_val = subprocess.check_output(command).decode('cp1252')
    assert 'Everything is Ok' in ret_val, f'something went wrong: {ret_val}'

    # TODO: parse ret_val into something useful
    return ret_val


def archive_extract(archive, into_dir=None, password=None, flat=False, overwrite=True, verbose=3):
    """
    extract an archive
    if a split-volumes archive, specify the first file (example.7z.001)
    :param archive: path to archive
    :param into_dir: where to extract to
    :param password: ascii only, excluding null bytes and double-quote char
    :param flat: extract all files into target dir ignoring directory structure
    :param overwrite: True or False or advanced options
    :param verbose: loudness as int {0, 1, 2, 3}
    :return: output printed by 7zip
    """
    archive_path = os.path.abspath(archive)
    assert os.path.isfile(archive_path), 'archive does not exist at provided path'

    # set arbitrary password if none given
    if password is None:
        password = '\x7f'  # ascii for DEL key

    # validity of provided password (if any)
    assert len(password) > 0, 'password must be at least one character long'

    # set arbitrary password if none given
    if into_dir is None:
        into_dir = '.'  # this directory

    # validity of provided dirname (if any)
    into_dir = os.path.abspath(into_dir.strip())
    assert len(into_dir) > 0, 'output dir name must be at least one character long'
    assert all(char not in os.path.basename(into_dir) for char in '\\/:*?"<>|'), 'invalid dir name'
    assert not os.path.isfile(into_dir), 'file exists at output dir location, cannot create folder'

    # need to write and run batch file to use this char
    if '"' in password:
        raise NotImplementedError('double-quote not supported')

    if overwrite is True:
        overwrite = 'a'
    elif overwrite is False:
        overwrite = 's'
    assert overwrite in {
        'a',  # overwrite without prompt
        's',  # skip
        'u',  # auto rename extracted file
        't',  # auto rename existing file
    }

    # make command
    command = [EXE_PATH,
               'e' if flat else 'x',  # decide which flag to use
               f'-o{into_dir}',
               f'-p{password}',
               f'-ao{overwrite}',
               f'-bb{verbose}',
               '-bt',
               archive_path,  # must be last argument for extraction
               ]

    # make the file, check that it's okay
    ret_val = subprocess.check_output(command).decode('cp1252')
    assert 'Everything is Ok' in ret_val, f'something went wrong: {ret_val}'

    # TODO: parse ret_val into something useful
    return ret_val


if __name__ == '__main__':
    import fnmatch
    import shutil


    def crawl(top, file_pattern='*'):
        for path, dir_list, file_list in os.walk(top):
            for file_name in fnmatch.filter(file_list, file_pattern):
                yield os.path.join(path, file_name)


    # cleanup existing 7z
    if os.path.exists('test.7z'):
        os.remove('test.7z')

    # cleanup existing folder
    if os.path.exists('test_output'):
        shutil.rmtree('test_output')

    # create archive
    print('CREATING ARCHIVE')
    print(archive_create(['cmds.txt', '__init__.py'],
                         'test.7z',
                         password='test password!',
                         encrypt_headers=True,
                         ))

    # test archive
    print('TESTING ARCHIVE')
    print(archive_test('test.7z',
                       password='test password!'))

    # extract archive
    print('EXTRACTING ARCHIVE')
    print(archive_extract('test.7z',
                          into_dir='test_output',
                          password='test password!'))
