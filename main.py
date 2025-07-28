import flet as ft
import flet_map
import file_actions as fa

class ProgressBanner(ft.Banner):
    def __init__(self, content=None, actions=None):
        super().__init__(content, actions)
        self.bgcolor = '#141414'
        self.leading = ft.Icon(ft.Icons.CHANGE_CIRCLE_OUTLINED, color=ft.Colors.WHITE, size=30)

        info_col = ft.Column(
            controls=[ft.Text(value="PoleTool is working. Please wait.", color=ft.Colors.WHITE), ft.ProgressBar(color=ft.Colors.BLUE)]
        )
        info = ft.Container(content=info_col)
        self.content = info
        self.actions = [ft.Container()]  # Actions must contain something. Using an empty container.


class FileAction(ft.Container):
    def __init__(self, name, description):
        super().__init__()
        act_switch = ft.Switch()
        self.title = ft.Text(value=name, color=ft.Colors.WHITE, size=16)
        action = ft.Row(controls=[act_switch, self.title], spacing=10)
        action_desc = ft.Text(description, color=ft.Colors.GREY_500, size=12)
        items = ft.Column(
            controls=[action, action_desc]
        )
        self.content = items
        self.bgcolor = '#383838'
        self.padding = 10
        self.border_radius = 15


class FileActions(ft.Column):
    def __init__(self):
        super().__init__()
        vetro = FileAction(name='Prepare for Vetro', description='Prepares information for Vetro import in order to automatically math up attributes when imported.')
        mrn = FileAction(name='Generate Make Ready Notes', description='Generates Make Ready Notes which are typically submitted with strand maps.')
        verizon = FileAction(name='Generate Verizon Application', description='Generates Verizon applications by municipality to be submitted for pole applications')
        frontier = FileAction(name='Generate Frontier Application', description='Generates Frontier applications to be submitted for pole applications')
        self.controls = [vetro, mrn, verizon, frontier]
        self.padding = 10
        self.scroll = ft.ScrollMode.ADAPTIVE

    def get_selected_actions(self):
        selected_actions = []
        for action in self.controls:
            if action.content.controls[0].controls[0].value:
                selected_actions.append(action.title.value)
        return selected_actions


class FileActionArea(ft.Container):
    def __init__(self):
        super().__init__()
        self.width = 400
        self.padding = 10
        self.actions_list = FileActions()
        items = ft.Column(
            controls=[
                ft.Text(value="File Actions", color=ft.Colors.WHITE, size=20),
                ft.Container(content=self.actions_list, expand=True)
            ]
        )
        self.content = items
        self.bgcolor = '#242424'


class PoleMap(flet_map.Map):
    def __init__(self, **kwargs):
        self.markers_ref = ft.Ref[flet_map.MarkerLayer]()
        layers = [
            flet_map.TileLayer(url_template="https://a.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}.png"),
            flet_map.MarkerLayer(ref=self.markers_ref, markers=[])
        ]
        super().__init__(
            layers=layers,
            initial_center=flet_map.MapLatitudeLongitude(39.8832154904524, -98.1021255445858),
            initial_zoom=3,
            expand=True,
            **kwargs
        )

        self.point_colors = {
            'PPL Company': ft.Colors.RED_700,
            'Verizon Pennsylvania Inc.': ft.Colors.GREEN_700,
            'Frontier Communications of PA. - New Holland': ft.Colors.PURPLE_700,
            'Frontier Communications of PA. - New Holland Telecom': ft.Colors.PURPLE_700,
            'Frontier Communications - Lakewood': ft.Colors.PURPLE_700,
            'Frontier Communications - Lakewood Telecom': ft.Colors.PURPLE_700,
            'Commonwealth Telephone Co.  dba Frontier Comm.': ft.Colors.PURPLE_700,
            'Commonwealth Telephone Co.  dba Frontier Comm. Telecom': ft.Colors.PURPLE_700,
            'Loop Telecom Pennsylvania LLC': ft.Colors.BLUE_700,
            'UGI Utilities - Electric Division': ft.Colors.ORANGE_700,
            'UGI Utilities - Gas': ft.Colors.ORANGE_700,
            'UGI PENN NATURAL GAS, INC': ft.Colors.ORANGE_700,
            'Service Electric Cablevision Inc - Mahanoy City': ft.Colors.BLUE_200,
            'Service Electric Cablevision': ft.Colors.BLUE_200,
            'Service Electric Cable TV Inc.': ft.Colors.BLUE_200,
            'Service Electric Company - Wilkes-Barre': ft.Colors.BLUE_200,
            'Upper Oxford Twp, Chester Co.': ft.Colors.TEAL_700,
            'City of Scranton - Wireless': ft.Colors.GREEN_200,
            'City of Scranton': ft.Colors.GREEN_200,
            'City of Scranton Office of Economic & Community Development': ft.Colors.GREEN_200,
            'CTSI, LLC, dba Frontier Communications': ft.Colors.AMBER_700
        }

    def create_point(self, latitude, longitude, company):
        point = flet_map.Marker(
            coordinates=flet_map.MapLatitudeLongitude(latitude, longitude),
            content=ft.Icon(ft.Icons.LOCATION_PIN, size=15, color=self.point_colors.get(company) or ft.Colors.GREY),
        )
        if self.markers_ref.current:
            self.markers_ref.current.markers.append(point)
            if self.page:
                self.page.update()


class FileIO(ft.Column):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.file_path = None
        self.output_path = None
        self.input_file = None
        self.file_picker = ft.FilePicker(on_result=self.pick_files_result)
        self.file_putter = ft.FilePicker(on_result=self.pick_dir_result)

        self.selected_file = ft.Text(value="No files selected", size=12)
        input_area = ft.Row(
            [
                ft.ElevatedButton(
                    "Choose Input File",
                    icon=ft.Icons.UPLOAD_FILE,
                    on_click=lambda _: self.file_picker.pick_files(allow_multiple=True),
                    icon_color=ft.Colors.GREEN,
                    color=ft.Colors.WHITE
                ),
                self.selected_file
            ]
        )

        self.selected_output_path = ft.Text(value="No output path selected", size=12)
        output_area = ft.Row(
            [
                ft.ElevatedButton(
                    "Choose Output Path",
                    icon=ft.Icons.DRIVE_FILE_MOVE_ROUNDED,
                    on_click=lambda _: self.file_putter.get_directory_path(),
                    icon_color=ft.Colors.YELLOW,
                    color=ft.Colors.WHITE
                ),
                self.selected_output_path
            ]
        )

        self.output_field = ft.TextField(label="Output File Name", hint_text="Enter a file name", expand=True, text_size=12)
        process_button = ft.ElevatedButton("Process", on_click=self.process_file, width=100, color=ft.Colors.PURPLE_300)
        process_area = ft.Row(
            [
                self.output_field,
                process_button
            ]
        )

        self.controls = [input_area, output_area, process_area]

    def pick_dir_result(self, e: ft.FilePickerResultEvent):
        if e.path:
            self.output_path = e.path
            self.selected_output_path.value = e.path
        else:
            self.selected_output_path.value = "Cancelled!"
        self.selected_output_path.update()

    def pick_files_result(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.file_path = e.files[0].path
        self.selected_file.value = (", ".join(map(lambda f: f.name, e.files)) if e.files else "Cancelled!")
        pole_points = fa.get_pole_points(self.file_path)
        for point in pole_points:
            map_and_file.map.create_point(point[0][0], point[0][1], point[1])
        self.selected_file.update()

    def process_file(self, _):
        if not self.file_path:
            self.selected_file.color = ft.Colors.RED_700
            self.update()
            return
        else:
            self.selected_file.color = ft.Colors.WHITE

        if not self.output_path:
            self.selected_output_path.color = ft.Colors.RED_700
            self.update()
            return
        else:
            self.selected_output_path.color = ft.Colors.WHITE

        if not self.output_field.value:
            self.output_field.border_color = ft.Colors.RED_700
            self.update()
            return
        else:
            self.output_field.border_color = ft.Colors.WHITE

        self.update()

        banner = ProgressBanner()
        self.page.open(banner)
        footer.status.value = 'Working...'
        self.page.update()
        file = fa.read_and_normalize(self.file_path)
        file_operations_dict = {
            'Prepare for Vetro': lambda: fa.vetro_export(file, self.output_path, self.output_field.value),
            'Generate Make Ready Notes': lambda: fa.generate_mrn(file, self.output_path, self.output_field.value),
            'Generate Verizon Application': lambda: fa.verizon_app(file, self.output_path, self.output_field.value),
            'Generate Frontier Application': lambda: fa.frontier_pdf(file, self.output_path, self.output_field.value)
        }
        operations = file_action.actions_list.get_selected_actions()
        for operation in operations:
            file_operations_dict[operation]()

        footer.status.value = 'Complete'
        footer.status.color = ft.Colors.GREEN
        self.page.close(banner)
        self.page.update()


class MapAndFileArea(ft.Container):
    def __init__(self):
        super().__init__()
        self.width = 600
        self.padding = 10
        self.file_io = FileIO()
        self.map = PoleMap()
        items = ft.Column([self.map, self.file_io])
        self.content = items
        self.bgcolor = '#383838'


class Footer(ft.Container):
    def __init__(self):
        super().__init__()
        self.status = ft.Text(value="Ready", color=ft.Colors.WHITE)
        self.version = ft.Text(value='PoleTool V2.0', color=ft.Colors.GREY_700)
        self.padding = 5
        self.bgcolor = '#141414'
        holder = ft.Row(controls=[self.version, self.status], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
        self.content = holder


file_action = FileActionArea()
map_and_file = MapAndFileArea()
footer = Footer()

def main(page):
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.window.width = 1015
    page.window.height = 700
    page.window.resizable = False
    page.title = "PoleTool v2.0 - by Greg Mocanu"
    page.bgcolor = '#242424'
    page.padding = 0
    page.spacing = 0

    page.overlay.append(map_and_file.file_io.file_picker)
    page.overlay.append(map_and_file.file_io.file_putter)

    areas = ft.Row(
        controls=[
            file_action,
            map_and_file
        ],
        expand=True,
        spacing=0
    )
    page.add(areas)
    page.add(footer)


ft.app(main)
