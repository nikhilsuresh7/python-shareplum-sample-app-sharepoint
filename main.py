from shareplum_utils import SharePoint

name = 'Clients'
description = 'Clients Description'

sample_add_list_data = [
    {'Title': 'Client 1'},
    {'Title': 'Client 5'},
    {'Title': 'Client 7'}
]

if __name__ == "__main__":
    # share point object
    sp_object = SharePoint()

    # List - Create
    sp_object.create_list(list_name=name, description=description)

    # List Items - New
    sp_object.add_update_list_items(list_name=name, data=sample_add_list_data, operation_kind='New')

    # Fetch and delete all list items
    all_list_data = sp_object.get_list_items(name, fields=['ID'])
    all_list_ids = [item.get('ID') for item in all_list_data]
    sp_object.delete_list_items(list_name=name, data=all_list_ids, operation_kind='Delete')

    # List - Delete
    # sp_object.delete_list(name)


