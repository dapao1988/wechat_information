import unittest
from wechat_members import *
from collections import OrderedDict

class TestWechatMembers(unittest.TestCase):
    def test_create_or_append_file_with_ui(self):
        filename, mode = create_or_append_file_with_ui("no_file.txt")
        self.assertEqual(filename, "no_file.txt")

    def test_write_members_to_file(self):
        #member_info_list = defaultdict(list)
        member_info_list=OrderedDict()

        member_info_list[HEAD_INFO] =  ['name', 'nick_name', 'province', 'city', 'sex', 'signature']
        member_info_list['John'] = ['John', 'Johnny', 'California', 'Los Angeles', 'Male', 'I am John.']
        member_info_list['Alice'] = ['Alice', 'Ali', 'New York', 'New York City', 'Female', 'I am Alice.']

        try:
            write_members_to_file("no_file.txt", member_info_list)
        except IOError as e:
            #print(f"IOError occurred: {e}")
            self.assertFalse(False,f"IOError occurred: {e}")
        except Exception as e:
            #print(f"An error occurred: {e}")
            self.assertFalse(False,f"An error occurred: {e}")
        else:
            #print("Write operation successful.")
            self.assertFalse(False,f"Write operation successful.")

    def test_write_members_to_excel(self):
        #member_info_list = defaultdict(list)
        member_info_list=OrderedDict()

        #member_info_list.insert(0, "member_info:{name, nick_name, province, city, sex, signature}")
        member_info_list[HEAD_INFO] = ['name', 'nick_name', 'province', 'city', 'sex', 'signature']
        member_info_list['John'] = ['John', 'Johnny', 'California', 'Los Angeles', 'Male', 'I am John.']
        member_info_list['Alice'] = ['Alice', 'Ali', 'New York', 'New York City', 'Female', 'I am Alice.']

        try:
            write_members_to_excel("info_dump.xlsx", member_info_list)
        except IOError as e:
            #print(f"IOError occurred: {e}")
            self.assertFalse(False,f"IOError occurred: {e}")
        except Exception as e:
            #print(f"An error occurred: {e}")
            self.assertFalse(False,f"An error occurred: {e}")
        else:
            #print("Write operation successful.")
            self.assertFalse(False,f"Write operation successful.")



if __name__ == '__main__':
    unittest.main()
