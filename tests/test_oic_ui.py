import unittest
from types import SimpleNamespace

from textual.widgets import DataTable, Input, Select

import oic_collector
import oic_query


class QueryDesignerUITests(unittest.IsolatedAsyncioTestCase):
    async def test_category_axes_shortcuts_and_search_focus(self):
        app = oic_collector.launch_terminal_ui(
            SimpleNamespace(codes=["MV"], timeout=60, retries=1), return_app=True
        )
        async with app.run_test(size=(150, 70)) as pilot:
            self.assertEqual(app.selected_code, "MV")
            self.assertTrue(app.axis_x)
            self.assertTrue(app.axis_y)
            table = app.query_one("#field_table", DataTable)
            self.assertGreater(table.row_count, 0)
            # Select one known hierarchy, move it with a keyboard shortcut.
            field = oic_query.PROFILES["MV"]["row1"]
            table.focus()
            await pilot.pause()
            app.selected_field = field
            await pilot.press("a")
            self.assertIn(field, app.axis_x)
            self.assertNotIn(field, app.axis_y)
            # Search Input must retain a/s as text, not trigger axis movement.
            before = list(app.axis_y)
            inp = app.query_one("#field_search", Input)
            inp.focus()
            await pilot.press("a", "s")
            self.assertIn("as", inp.value)
            self.assertEqual(app.axis_y, before)
            # Category switch resets the view to that cube's profile.
            app.select_category("FI")
            self.assertEqual(app.selected_code, "FI")
            self.assertEqual(app.axis_x, [oic_query.PROFILES["FI"]["year"]])
            self.assertEqual(app.axis_y, [oic_query.PROFILES["FI"]["row1"]])
            # Preview should not send requests.
            app.preview_selected_query()
            self.assertIn("FI_AGGREGATE_FY", app.query_one("#preview_text").render().plain
                          if hasattr(app.query_one("#preview_text").render(), "plain")
                          else str(app.query_one("#preview_text").render()))


if __name__ == "__main__":
    unittest.main()
