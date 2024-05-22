import os

from textual.app import App
from textual.binding import Binding
from textual.containers import Container
from textual.widgets import Button, Input, Label, Header, Footer, TextArea

from script import transform_excel


class ExcelTransformerApp(App):
    CSS_PATH = "script-ui.tcss"
    TITLE = "Excel transformer script"
    BINDINGS = [Binding(key="Ctrl + C", action="quit", description="Quit the app")]

    def compose(self):
        yield Header()
        yield Label("Input Excel file:")
        yield Input(placeholder="Enter the path to the input Excel file (example: input.xlsx)", id="input_file")
        yield Label("Output Excel file:")
        yield Input(placeholder="Enter the path to the output Excel file (example: output.xlsx)", id="output_file")
        yield Label("Start column name:")
        yield Input(placeholder="Enter the name of the first column to modify (example: First column to modify)", id="start_column")
        yield Button(label="Transform", id="transform_button", variant="primary")
        yield Container(id="status_text_container")
        yield Footer()

    def on_button_pressed(self, event: Button.Pressed) -> None:
        """Event handler called when a button is pressed."""
        if event.button.id == "transform_button":
            self.transform_excel()
        elif event.button.id == "open_output_excel_button":
            self.open_output_excel()

    def transform_excel(self):
        # Extract input values
        input_file = self.query_one("#input_file", Input).value
        output_file = self.query_one("#output_file", Input).value
        start_column_name = self.query_one("#start_column", Input).value

        # Perform the transformation
        try:
            transform_excel(input_file, output_file, start_column_name)

            self.query_one("#status_text_container").remove_children()
            self.query_one("#status_text_container").mount(Label("✅ Successfully transformed Excel file", id="success_text"))
            self.query_one("#status_text_container").mount(Button(label="Open output Excel file", id="open_output_excel_button", variant="success"))
        except Exception as e:
            self.query_one("#status_text_container").remove_children()
            self.query_one("#status_text_container").mount(
                TextArea(f"Error: {str(e)}", read_only=True, id="error_text"))  # text area to achieve text wrapping

    def open_output_excel(self):
        output_file = self.query_one("#output_file", Input).value
        # Open the file in the OS (this open command only works on MacOs)
        os.system(f"open {output_file}")


if __name__ == "__main__":
    ExcelTransformerApp().run()
