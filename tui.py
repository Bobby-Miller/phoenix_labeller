from pathlib import Path
from textual.app import App, ComposeResult
from textual.widgets import Header, Footer, Button, Label, Log, Input
from textual.containers import Vertical, Horizontal
from textual_fspicker import FileOpen, SelectDirectory, Filters

from main import generate_mtp_solution

script_dir = Path(__file__).resolve().parent

MAT_FILE_LOCATION = script_dir / "TR_WML6(13X13)R.mat"
class MTPGeneratorTUI(App):
    CSS = """
    Vertical { padding: 1; }
    Label { margin-left: 2; color: $text-muted; min-width: 20; }
    Input { width: 1fr; margin-left: 2; }
    Log {
            height: 1fr;   /* THIS IS KEY: Takes only the REmaining space */
            min-height: 5;
            border: tall $accent;
            background: $boost;
            margin: 1;
        }
    .btn-column { width: 25; }
    """

    BINDINGS = [("q", "quit", "Quit")]

    def __init__(self):
        super().__init__()
        self.source_xlsx = None
        self.output_dir = None

    def compose(self) -> ComposeResult:
        yield Header()
        with Vertical():
            # Source XLSX
            with Horizontal():
                yield Button("Source XLSX", id="select_xlsx", classes="btn-column")
                yield Label("Not selected", id="lbl_xlsx")
            
            # Output Directory
            with Horizontal():
                yield Button("Output Directory", id="select_dir", classes="btn-column")
                yield Label("Not selected", id="lbl_dir")

            # Filename Input
            with Horizontal():
                yield Label("Output Filename:", classes="btn-column")
                yield Input(placeholder="filename.mtp", id="filename_input")

            yield Button("Generate MTP", id="run", variant="success")
            yield Log(id="status_log")
        yield Footer()

    def on_button_pressed(self, event: Button.Pressed) -> None:
        log = self.query_one("#status_log", Log)

        if event.button.id == "select_xlsx":
            filters = Filters(("Excel Files", lambda p: p.suffix.lower() == ".xlsx"), ("All", lambda _: True))
            self.push_screen(FileOpen(".", filters=filters), self.set_xlsx)
        
        elif event.button.id == "select_dir":
            self.push_screen(SelectDirectory(".", select_button="Select",
                                             cancel_button="Cancel"), self.set_dir)

        elif event.button.id == "run":
            filename = self.query_one("#filename_input", Input).value.strip()
            
            if all([self.source_xlsx, self.output_dir, filename]):
                # Construct the final path
                full_output_path = self.output_dir / filename
                
                try:
                    # Logic execution
                    generate_mtp_solution(str(self.source_xlsx), str(MAT_FILE_LOCATION), str(full_output_path))
                    log.write_line(f"SUCCESS: Created {full_output_path}")
                except Exception as e:
                    log.write_line(f"Error on logic execution")
                    log.write_line(f"ERROR: {e}")
            else:
                log.write_line("MISSING DATA: Ensure all files, directory, and filename are set.")

    def set_xlsx(self, path: Path) -> None:
        if path:
            self.source_xlsx = path
            self.query_one("#lbl_xlsx", Label).update(path.name)


    def set_dir(self, path: Path) -> None:
        if path:
            self.output_dir = path
            self.query_one("#lbl_dir", Label).update(f"DIR: {path.name}/")

if __name__ == "__main__":
    app = MTPGeneratorTUI()
    app.run()

