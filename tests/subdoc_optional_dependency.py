import importlib.abc
import sys


class BlockDocxcompose(importlib.abc.MetaPathFinder):
    missing_name = "docxcompose"

    def find_spec(self, fullname, path, target=None):
        if fullname == "docxcompose" or fullname.startswith("docxcompose."):
            raise ModuleNotFoundError(
                "No module named '%s'" % self.missing_name,
                name=self.missing_name,
            )
        return None


blocker = BlockDocxcompose()
sys.meta_path.insert(0, blocker)

try:
    from docxtpl import DocxTemplate

    template = DocxTemplate("templates/subdoc_tpl.docx")

    try:
        template.new_subdoc()
    except ModuleNotFoundError as exc:
        assert exc.name == "docxcompose"
        assert "new_subdoc() requires the optional docxcompose dependency" in str(exc)
        assert 'pip install "docxtpl[subdoc]"' in str(exc)
        assert isinstance(exc.__cause__, ModuleNotFoundError)
    else:
        raise AssertionError("new_subdoc() did not report the missing dependency")

    blocker.missing_name = "some_other_dependency"
    try:
        template.new_subdoc()
    except ModuleNotFoundError as exc:
        assert exc.name == "some_other_dependency"
        assert str(exc) == "No module named 'some_other_dependency'"
        assert "pip install" not in str(exc)
    else:
        raise AssertionError("an unrelated import failure was masked")
finally:
    sys.meta_path.remove(blocker)
