import unittest
from directory_search_base import directory_search


class MyTestCase(unittest.TestCase):
    def test_something(self):
        pseudo_subfolder_1 =["TESTINVOICE1-V01", "TESTINVOICE1-V02", "TESTINVOICE1-V03", "TESTINVOICE1-V04"]
        pseudo_subfolder_2 =["TESTINVOICE1-V05", "TESTINVOICE1-V06", "TESTINVOICE1-V07", "TESTINVOICE1-V10"]
        pseudo_subfolder_3 =["TESTINVOICE1-V12", "TESTINVOICE1-V13", "TESTINVOICE14", "TESTINVOICE1-V15"]
        pseudo_subfolder_4 =["TESTINVOICE1-V17", "TESTINVOICE1-V18", "TESTINVOICE1-V19"]
        pseudo_subfolder_5 =[]
        pseudo_main_folder =[pseudo_subfolder_1, pseudo_subfolder_2, pseudo_subfolder_3, pseudo_subfolder_4, pseudo_subfolder_5]
        found_files =['','','','','','']
        files_to_find =["TESTINVOICE1-V05", "TESTINVOICE1-V07", "TESTINVOICE1-V11", "TESTINVOICE1-V17", "TESTINVOICE1-V04"]
        for p_folder in pseudo_main_folder:
            for p_file in p_folder:
                if p_file in files_to_find:
                    position = files_to_find.index(p_file)
                    found_files.insert(position, p_file)

        result = found_files
        print(result)
        self.assertEqual(result[2], "TESTINVOICE1-V17")



if __name__ == '__main__':
    unittest.main()
