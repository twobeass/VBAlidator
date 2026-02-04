import unittest
from src.lexer import Lexer
from src.parser import VBAParser

class TestSingleLineBlocks(unittest.TestCase):
    def test_single_line_type(self):
        code = """
        Attribute VB_Name = "TestMod"
        Type Point : x As Long : y As Long : End Type
        """
        lexer = Lexer(code)
        tokens = list(lexer.tokenize())
        parser = VBAParser(tokens)
        module = parser.parse_module()

        self.assertIn("Point", module.types)
        point_type = module.types["Point"]
        self.assertEqual(len(point_type.members), 2)
        self.assertEqual(point_type.members[0].name, "x")
        self.assertEqual(point_type.members[1].name, "y")

    def test_single_line_type_mixed(self):
        code = """
        Attribute VB_Name = "TestMod"
        Type Rect
            top As Long : left As Long
            bottom As Long : right As Long
        End Type
        """
        lexer = Lexer(code)
        tokens = list(lexer.tokenize())
        parser = VBAParser(tokens)
        module = parser.parse_module()

        self.assertIn("Rect", module.types)
        rect_type = module.types["Rect"]
        self.assertEqual(len(rect_type.members), 4)
        self.assertEqual(rect_type.members[1].name, "left")
        self.assertEqual(rect_type.members[2].name, "bottom")

    def test_single_line_enum(self):
        code = """
        Attribute VB_Name = "TestMod"
        Enum Colors : Red = 1 : Green = 2 : Blue = 3 : End Enum
        """
        lexer = Lexer(code)
        tokens = list(lexer.tokenize())
        parser = VBAParser(tokens)
        module = parser.parse_module()

        self.assertIn("Colors", module.types)
        colors_enum = module.types["Colors"]
        self.assertEqual(len(colors_enum.members), 3)
        self.assertEqual(colors_enum.members[0].name, "Red")
        self.assertEqual(colors_enum.members[1].name, "Green")
        self.assertEqual(colors_enum.members[2].name, "Blue")

    def test_single_line_enum_mixed(self):
        code = """
        Attribute VB_Name = "TestMod"
        Enum Status
            Ready = 0 : Busy = 1
            Error = 2
        End Enum
        """
        lexer = Lexer(code)
        tokens = list(lexer.tokenize())
        parser = VBAParser(tokens)
        module = parser.parse_module()

        self.assertIn("Status", module.types)
        status_enum = module.types["Status"]
        self.assertEqual(len(status_enum.members), 3)
        self.assertEqual(status_enum.members[1].name, "Busy")
        self.assertEqual(status_enum.members[2].name, "Error")

    def test_enum_value_skipping_with_colon(self):
        # Ensure that parsing 'A = 1 : B = 2' doesn't consume 'B = 2' as part of A's value
        code = """
        Attribute VB_Name = "TestMod"
        Enum TestEnum : A = 1 : B = 2 : End Enum
        """
        lexer = Lexer(code)
        tokens = list(lexer.tokenize())
        parser = VBAParser(tokens)
        module = parser.parse_module()

        enum_obj = module.types["TestEnum"]
        names = [m.name for m in enum_obj.members]
        self.assertIn("A", names)
        self.assertIn("B", names)

if __name__ == '__main__':
    unittest.main()
