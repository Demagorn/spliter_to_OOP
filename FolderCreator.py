import pathlib

class Folder:
    def __init__(self, destination_path):
        """checks if folder exist if not creates - no exceptions raised"""
        self.destination = destination_path
        self.create_folder()

    def create_folder(self):
        pathlib.Path(self.destination).mkdir(parents=True, exist_ok=True)