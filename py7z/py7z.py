import os
import subprocess
import sys

__author__ = 'Avery'

SYS_BITS = sys.maxsize.bit_length() + 1  # 32-bit or 64-bit
assert SYS_BITS in (32, 64)

EXE_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), u'7-zip/x%d/7z.exe' % SYS_BITS))
assert os.path.exists(EXE_PATH)


def archive_create(files_and_folders, archive, password=None, encrypt_headers=False, overwrite=False, verbose=3,
                   volumes=None):
    """
    create 7z archive from some files and folders
    files and folders will be placed in the root of the archive

    :param volumes: size of volume, eg '10k' or '2g'
    :param verbose: loudness as int {0, 1, 2, 3}
    :param overwrite: True or False
    :param files_and_folders: list of paths
    :param archive: path (or name)
    :param password: ascii only, excluding null bytes and double-quote char
    :param encrypt_headers: encrypt file names and directory tree within the archive
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
            print(u'multiple items with same filename added to archive: {NAME}\n{PATHS}'.format(
                NAME=item_name,
                PATHS='\n'.join('<{PATH}>'.format(PATH=item_path) for item_path in item_paths),
            ), file=sys.stderr)
            raise ValueError(item_name)

    # 7z will not create an archive over an existing dir
    assert not os.path.isdir(archive_path), u'dir already exists at output path'

    # don't want to accidentally tell 7zip to overwrite anything
    assert overwrite or not os.path.isfile(archive_path), u'file already exists at output path'

    # base command (without input files/folders)
    command = [EXE_PATH,
               u'a',
               archive_path,
               u'-t7z',
               u'-m0=lzma2',
               u'-mx=9',
               u'-bb{VERBOSE}'.format(VERBOSE=verbose),
               u'-bt',
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
            command += [u'-v{SIZE}'.format(SIZE=size)]

    # add password and header encryption
    if password is not None:
        assert len(password) > 0, u'password must be at least one character long'

        # need to write and run batch file to use this char
        if u'"' in password:
            raise NotImplementedError(u'double-quote not supported')

        # command += u' "-p{PASSWORD}"'.format(PASSWORD=password)
        command += [u'-p{PASSWORD}'.format(PASSWORD=password)]
        if encrypt_headers:
            command += [u'-mhe']

    # add files
    command += todo_paths

    # make the file, check that it's okay
    ret_val = subprocess.check_output(command).decode('cp1252')
    assert u'Everything is Ok' in ret_val, u'something went wrong: ' + ret_val

    # todo: parse ret_val into something useful
    return ret_val


def archive_test(archive, password=None, verbose=3):
    """
    test the integrity of an archive

    :param verbose: loudness as int {0, 1, 2, 3}
    :param archive: path (or name)
    :param password: ascii only, excluding null bytes and double-quote char
    :return: output printed by 7zip
    """
    archive_path = os.path.abspath(archive)

    # set arbitrary password if none given
    if password is None:
        password = u'\x7f'  # ascii for DEL key

    # validity of provided password (if any)
    assert len(password) > 0, u'password must be at least one character long'

    # need to write and run batch file to use this char
    if u'"' in password:
        raise NotImplementedError(u'double-quote not supported')

    # always supply a password
    command = [EXE_PATH,
               u't',
               archive_path,
               u'-p{PASSWORD}'.format(PASSWORD=password),
               u'-bb{VERBOSE}'.format(VERBOSE=verbose),
               u'-bt',
               ]

    # make the file, check that it's okay
    ret_val = subprocess.check_output(command).decode('cp1252')
    assert u'Everything is Ok' in ret_val, u'something went wrong: ' + ret_val

    # todo: parse ret_val into something useful
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
    assert os.path.isfile(archive_path), u'archive does not exist at provided path'

    # set arbitrary password if none given
    if password is None:
        password = u'\x7f'  # ascii for DEL key

    # validity of provided password (if any)
    assert len(password) > 0, u'password must be at least one character long'

    # set arbitrary password if none given
    if into_dir is None:
        into_dir = u'.'  # this directory

    # validity of provided dirname (if any)
    into_dir = os.path.abspath(into_dir.strip())
    assert len(into_dir) > 0, u'containing dir name must be at least one character long'
    assert all(char not in os.path.basename(into_dir) for char in u'\\/:*?"<>|'), u'invalid dir name'
    assert not os.path.isfile(into_dir), u'dir already exists at output path, cannot create folder'

    # need to write and run batch file to use this char
    if u'"' in password:
        raise NotImplementedError(u'double-quote not supported')

    if overwrite is True:
        overwrite = 'a'
    elif overwrite is False:
        overwrite = 's'
    assert overwrite in [
        'a',  # overwrite without prompt
        's',  # skip
        'u',  # auto rename extracted file
        't',  # auto rename existing file
    ]

    # make command
    command = [EXE_PATH,
               u'xe'[flat],  # decide which flag to use
               u'-o{DIR}'.format(DIR=into_dir),
               u'-p{PASSWORD}'.format(PASSWORD=password),
               u'-ao{OVERWRITE}'.format(OVERWRITE=overwrite),
               u'-bb{VERBOSE}'.format(VERBOSE=verbose),
               u'-bt',
               archive_path,  # must be last argument for extraction
               ]

    # make the file, check that it's okay
    ret_val = subprocess.check_output(command).decode('cp1252')
    assert u'Everything is Ok' in ret_val, u'something went wrong: ' + ret_val

    # todo: parse ret_val into something useful
    return ret_val


if __name__ == '__main__':
    import fnmatch


    def crawl(top, file_pattern='*'):
        for path, dir_list, file_list in os.walk(top):
            for file_name in fnmatch.filter(file_list, file_pattern):
                yield os.path.join(path, file_name)


    # cleanup
    if os.path.exists(r'passwd.7z'):
        os.remove(r'passwd.7z')

    # create archive
    print(archive_create([r'cmds.txt', r'__init__.py'],
                         r'passwd.7z',
                         password=('passwd'),
                         encrypt_headers=True,
                         ))

    # test archive
    print(archive_test(r'passwd.7z',
                       password='passwd'))

    # extract archive
    print(archive_extract(r'passwd.7z',
                          into_dir=r'passwd',
                          password='passwd'))
