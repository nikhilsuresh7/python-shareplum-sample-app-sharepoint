from shareplum_utils import SharePoint

name = 'Clients'
description = 'Clients Description'


if __name__ == "__main__":
    # share point object
    sp_object = SharePoint()

    # List - Create
    sp_object.create_list(list_name=name, description=description)

    # List - Delete
    # sp_object.delete_list(name)
