import os
import sys
import argparse
import PyPDF2 as Pdf
import tkinter as tk
import functools
import multiprocessing
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from typing import Union


class ProgressBar:
    def __init__(self, parent: tk.Tk = None, length=300):
        self.parent = parent
        if self.parent is None:
            self.parent = tk.Tk()

        self.parent.resizable(width=False, height=False)
        self.parent.title("PDFMerger")
        self.frame = ttk.Frame(self.parent, padding=10, relief='sunken')
        self.frame.grid(row=0, column=0, sticky='nsew')
        self.parent.grid_columnconfigure(0, weight=1)
        self.parent.grid_rowconfigure(0, weight=1)
        self.label = ttk.Label(
            self.frame,
            text="Merging files, please wait..."
        )
        self.progressbar = ttk.Progressbar(
            self.frame,
            orient='horizontal',
            mode='indeterminate',
            length=length
        )
        self.label.grid(row=0, column=0)
        self.progressbar.grid(row=1, column=0)
        self.progressbar.start(interval=15)

        def disable_event():
            pass

        # Prevent progressbar from being destroyed by the user before the
        # parent process terminates.
        self.parent.protocol("WM_DELETE_WINDOW", disable_event)
        self.parent.mainloop()

    def stop(self):
        """Destroy the progressbar."""
        self.parent.destroy()


def create_progressbar():
    ProgressBar()


def merge_pdfs(src: Union[tuple, list], dst: str, bookmarks=True):
    """
    Merge PDF files in src into a single pdf file (dst).

    :param src: path names of pdf files to merge
    :param dst: path name of output file
    :param bookmarks: flag specifying whether or not to import bookmarks, defaults to True
    """
    pbar_process = multiprocessing.Process(target=create_progressbar, daemon=True)
    pbar_process.start()
    try:
        file_objects = [Pdf.PdfReader(open(file, 'rb'), strict=False) for file in src]
        merger = Pdf.PdfMerger()
        for file_object in file_objects:
            merger.append(file_object, import_outline=bookmarks)

        merger.write(dst)
        merger.close()

    except Exception:
        # The progressbar will continue to run after an error occurs during merging
        # resulting in the error dialog window and progressbar being displayed at
        # the same time. To avoid this, the progressbar is destroyed if any exception
        # is thrown by merge_pdfs()
        pbar_process.terminate()
        raise


def ask_save_as(dir_=None):
    dst = filedialog.asksaveasfilename(
        initialdir=dir_,
        defaultextension=[('PDF', '*.pdf')],
        filetypes=[('PDF', '*.pdf')]
    )
    # Terminate program if no selection is made.
    if not dst:
        sys.exit(0)

    return dst


def ask_open_files(dir_=None):
    src = filedialog.askopenfilenames(
        title="select pdf files to merge",
        filetypes=[('PDF', '*.pdf')],
        initialdir=dir_
    )
    # Terminate program if no selection is made.
    if not src:
        sys.exit(0)

    return src


def main():
    # Fetch the commandline arguments passed, if any.
    parser = argparse.ArgumentParser(description="command line tool for merging PDF files")
    parser.add_argument(
        "src",
        type=str,
        help='full path names of files to merge',
        nargs='*',
        metavar='S',
        action='store',
        default=None
        )
    parser.add_argument(
        "--dst",
        "-d",
        type=str,
        help='full path name of merged file',
        metavar='D',
        default=None
        )
    parser.add_argument(
        "--bookmarks",
        "-b",
        action="store_false",
        default=True,
        help="whether or not to keep bookmarks from src files, defaults to True"
    )
    args = parser.parse_args()
    merge = functools.partial(merge_pdfs, bookmarks=args.bookmarks)
    # Merge src files into dst if both flags are specified.
    # If one or both are unspecified, create a dialog prompt for selecting src files
    # followed by a file 'save as' dialog.
    if args.src and args.dst:
        merge(args.src, args.dst)

    else:
        # Root widget is created and displayed by default when creating a dialog.
        # It is necessary to explicitly create and withdraw the root window to keep it
        # out of sight when the dialog pops up.
        root = tk.Tk()
        root.withdraw()
        if not args.src and not args.dst:
            src = ask_open_files()
            dst = ask_save_as()
            merge(src, dst)

        elif args.src and not args.dst:
            src_dir = os.path.split(args.src[0])[0]
            dst = ask_save_as(dir_=src_dir)
            merge(args.src, dst)

        elif args.dst and not args.src:
            dst_dir = os.path.split(args.dst[0])[0]
            src = ask_open_files(dir_=dst_dir)
            merge(src, args.dst)

        root.destroy()


if __name__ == "__main__":
    try:
        main()

    except Exception as e:
        _root = tk.Tk()
        _root.withdraw()
        messagebox.showerror(
            title="PDFMergerError",
            message="An error occurred when merging the selected files:",
            detail=e,
            icon="warning"
        )
        _root.destroy()
        raise
