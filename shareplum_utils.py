from shareplum import Site, Office365
from shareplum.site import Version

from typing import List, Dict, Union
from dotenv import dotenv_values

env_data = dotenv_values(".env")
missing_vars = set(dotenv_values(".env-example")) - set(env_data)

if missing_vars:
    print("Please provide a valid .env file and try again!")
    print(f"Missing - {missing_vars}")
    exit()


class SharePoint:

    def __init__(self):
        """
        Initializes and create a site object on successful authentication
        """
        try:
            self.authcookie = Office365(
                env_data.get('SHARE_POINT_URL'),
                username=env_data.get('USER_NAME'),
                password=env_data.get('PASSWORD')
                ).GetCookies()

            self.site = Site(
                env_data.get('SITE_URL'),
                version=Version.v365,
                authcookie=self.authcookie
                )

        except Exception as error:
            print(error)
            exit()
    
    def list_exists(self, list_name: str) -> bool:
        """
        Custom function to check if a list exists
        """
        all_lists = self.site.GetListCollection()
        list_titles = [item.get('Title') for item in all_lists]
        return True if list_name in list_titles else False

    def create_list(self, list_name: str, description:str, template_id=100) -> None:
        """
        Create a new list if it doesn't exists
        """
        if self.list_exists(list_name):
            print(f"List '{list_name}' already exists.")
        else:
            self.site.AddList(
                list_name,
                description=description,
                template_id=template_id
                )
            print(f"Created list '{list_name}'.")

    def delete_list(self, list_name:str) -> None:
        """
        Delete list if it exists
        """
        if self.list_exists(list_name):
            self.site.DeleteList(list_name)
            print(f"Deleted list '{list_name}'.")
        else:
            print(f"List '{list_name}' does not exists.")

    def add_update_list_items(self, list_name, data: List[Dict], operation_kind: str) -> None:
        """
        Add or update list items based on the value of 'operation_kind'
        """
        if self.list_exists(list_name):
            list_obj = self.site.List(list_name)

            if operation_kind not in ('New', 'Update'):
                print(f"Invalid Kind!")
                exit()

            list_obj.UpdateListItems(data=data, kind=operation_kind)
            print(f"Success - {operation_kind} list.")
        else:
            print(f"List '{list_name}' does not exists.")

    def delete_list_items(self, list_name, data: List[str], operation_kind: str or int) -> None:
        """
        Delete list items based on the value of 'operation_kind'
        """
        if self.list_exists(list_name):
            list_obj = self.site.List(list_name)

            if operation_kind != 'Delete':
                print(f"Invalid Kind!")
                exit()

            list_obj.UpdateListItems(data=data, kind=operation_kind)
            print(f"Success - {operation_kind} list.")
        else:
            print(f"List '{list_name}' does not exists.")

    def get_list_items(self, list_name:str, fields: List[str]) -> List[Dict]:
        """
        Get items list filtered by field names
        """
        if self.list_exists(list_name):
            list_obj = self.site.List(list_name)
            sp_data = list_obj.GetListItems(fields=fields)
            return sp_data














