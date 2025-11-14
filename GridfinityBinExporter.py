import itertools
import math
import platform
import shutil
import subprocess
import sys
import os

# include __pypackages__ / bin
sys.path.append(os.path.join(os.path.dirname(__file__), '__pypackages__'))
sys.path.append(os.path.join(os.path.dirname(__file__), 'bin'))
os.environ['PATH'] += f"{os.pathsep}{os.path.normpath(os.path.join(os.path.dirname(__file__), 'bin'))}"

import adsk.core, adsk.fusion, traceback
from typing import List, Literal
from timeit import default_timer as timer
from datetime import timedelta, datetime
import imageio
import pygifsicle
import zipfile, glob, re

class IDS:
    SLIDER_WALL = 'wall_thickness_slider'
    SLIDER_DIVISION = 'division_slider'
    BTN_EXPORT = 'btn-export'

    CBOX_GIF_ALL = 'cbox-gif-all'
    CBOX_GIF_Z = 'cbox-gif-z'
    CBOX_CREATE_IMAGE = 'cbox-create-images'

class PARAMS:
    MAGNET_DIAMETER = 'MagnetDiameter'
    MAGNET_REMOVE_DIAMETER = 'MagnetRemoveDiameter'
    MAGNET_DEPTH = 'MagnetDepth'

    WALL_THICKNESS = 'WallThickness'
    DIVISIONS = 'Divisions'

    SCOOP_RADIUS = "ScoopCurveRadius"

    X = 'Width_X'
    Y = 'Width_Y'
    Z = 'Height'

class INPUTS:
    grid_x: adsk.core.IntegerSliderCommandInput 
    grid_y: adsk.core.IntegerSliderCommandInput
    grid_z: adsk.core.IntegerSliderCommandInput
    grid_z_step: adsk.core.IntegerSliderCommandInput

    wall_thickness: List[adsk.core.FloatSpinnerCommandInput] = []
    division: adsk.core.IntegerSliderCommandInput

    scoop_radius: adsk.core.IntegerSliderCommandInput

    mag_diameter: adsk.core.FloatSpinnerCommandInput
    mag_rem_diameter: adsk.core.FloatSpinnerCommandInput
    mag_depth: adsk.core.FloatSpinnerCommandInput

    cbox_useless: adsk.core.BoolValueCommandInput
    cbox_create_images: adsk.core.BoolValueCommandInput
    cbox_gif_all: adsk.core.BoolValueCommandInput
    cbox_gif_z: adsk.core.BoolValueCommandInput
    skip_existing: adsk.core.BoolValueCommandInput
    zip: adsk.core.BoolValueCommandInput

    max_frames_per_gif: adsk.core.IntegerSpinnerCommandInput

    gif_fps: adsk.core.IntegerSpinnerCommandInput
    gif_lossy: adsk.core.IntegerSpinnerCommandInput
    gif_optimize: adsk.core.IntegerSpinnerCommandInput
    gif_colors: adsk.core.IntegerSpinnerCommandInput

    def clear_wall_thickness(self):
        self.wall_thickness = []

# --- Globals

G_APP = None
G_UI  = None
# Global set of event handlers to keep them referenced for the duration of the command
G_HANDLERS = []
G_INPUTS = INPUTS()

COPY_UPLOAD_WORTHY_STLS = False # If True, this will copy all 6x6x* files into a folder + no ZIP processing 

# --- INIT

INIT_X_END=10
INIT_Y_END=10
INIT_Z_END=18
INIT_Z_STEPS=3
INIT_DEFAULT_EXPORT_PATH = 'C:/Export-F360'

INIT_SCREENSHOT_W=640 # 1920 
INIT_SCREENSHOT_H=360 # 1080

# --- Templates

TPL_VARIANT_FOLDER = "{folder}/wall-{wall_width}/divisions-{divisions}"
TPL_VARIANT_NAME = "gfbin1.2_{x}x{y}x{z}_w{wall_width}d{divisions}"

# ---

class GridfinityBinExporter:
    __exporting = False
    __progress_dialog: adsk.core.ProgressDialog | None = None

    __export_folder = ''

    __screenshot_filenames: List[str] = []
    __screenshot_z_filenames: List[List[str]] = []

    __generate_gif_all = False
    __generate_gif_row = False

    __z_start: int
    __z_increment: int
    __range_x: range
    __range_y: range
    __range_z: range
    __range_div: range
    __generate_no_useless: bool
    __skip_existing_stl: bool
    __skip_image_creation: bool

    __list_ww: list

    __amount = 0
    __skipped = 0

    __design: adsk.fusion.Design
    __export_manager: adsk.fusion.ExportManager
    __bin_parameter_list: List[adsk.fusion.Parameter]

    def get_total_processed_stl(self):
        return self.__amount + self.__skipped

    def is_exporting(self):
        return self.__exporting

    def was_cancelled(self):
        return self.__progress_dialog.wasCancelled if self.__progress_dialog else True

    def stop_exporting(self):
        if self.__progress_dialog:
            self.__progress_dialog.hide()

        self.__exporting = False
        adsk.doEvents()

    def get_screenshot_folder(self):
        return f"{self.__export_folder}/screenshots"

    def calc_z(self, z_index: int):
        return self.__z_start + self.__z_increment * z_index

    def setup_export_folder(self) -> bool:
        folder_input = G_UI.inputBox('Input path to save folder: ', 'Define root of export', INIT_DEFAULT_EXPORT_PATH)
        if folder_input[1] or len(folder_input[0]) == 0:
            return False
        
        base_folder = f"{folder_input[0]}/bin_{G_INPUTS.grid_x.valueOne}-{G_INPUTS.grid_x.valueTwo}"
        base_folder += f"x{G_INPUTS.grid_y.valueOne}-{G_INPUTS.grid_y.valueTwo}"
        base_folder += f"x{G_INPUTS.grid_z.valueOne}-{G_INPUTS.grid_z.valueTwo}s{G_INPUTS.grid_z_step.valueOne}"
        base_folder += f"_d{G_INPUTS.division.valueOne}-{G_INPUTS.division.valueTwo}"

        self.__export_folder = base_folder

        os.makedirs(base_folder, exist_ok=True)
        os.makedirs(self.get_screenshot_folder(), exist_ok=True)

        return True
    
    def setup_fusion_params(self, design: adsk.fusion.Design):
        param_mag_diameter = design.allParameters.itemByName(PARAMS.MAGNET_DIAMETER)
        param_mag_rem_diameter = design.allParameters.itemByName(PARAMS.MAGNET_REMOVE_DIAMETER)
        param_mag_depth = design.allParameters.itemByName(PARAMS.MAGNET_DEPTH)
        param_scoop_radius = design.allParameters.itemByName(PARAMS.SCOOP_RADIUS)

        design.modifyParameters(
            [param_mag_diameter, param_mag_rem_diameter, param_mag_depth,param_scoop_radius],
            [
                adsk.core.ValueInput.createByString(G_INPUTS.mag_diameter.expression),
                adsk.core.ValueInput.createByString(G_INPUTS.mag_rem_diameter.expression),
                adsk.core.ValueInput.createByString(G_INPUTS.mag_depth.expression),
                adsk.core.ValueInput.createByString(G_INPUTS.scoop_radius.expressionOne)
            ]
        )

    def setup_ui_params(self):
        self.__skip_image_creation = G_INPUTS.cbox_create_images.value == False
        self.__generate_gif_all = G_INPUTS.cbox_gif_all.value and not self.__skip_image_creation
        self.__generate_gif_row = G_INPUTS.cbox_gif_z.value and not self.__skip_image_creation

        self.__z_start = G_INPUTS.grid_z.valueOne
        self.__z_increment = G_INPUTS.grid_z_step.valueOne

        self.__range_x = range(G_INPUTS.grid_x.valueOne, G_INPUTS.grid_x.valueTwo + 1)
        self.__range_y = range(G_INPUTS.grid_y.valueOne, G_INPUTS.grid_y.valueTwo + 1)
        self.__range_z = list(map(lambda z: self.calc_z(z), range(math.floor((G_INPUTS.grid_z.valueTwo - self.__z_start) / self.__z_increment) + 1)))
        self.__range_div = range(G_INPUTS.division.valueOne, G_INPUTS.division.valueTwo + 1)
        self.__generate_no_useless = G_INPUTS.cbox_useless.value
        self.__skip_existing_stl = G_INPUTS.skip_existing.value

        self.__list_ww = list(map(lambda input: self.__cm_into_mm(input.value), G_INPUTS.wall_thickness))

        # calc the amount of stl files
        amount_divisions = G_INPUTS.division.valueTwo + 1 - G_INPUTS.division.valueOne
        self.__progress_dialog.maximumValue = ((G_INPUTS.grid_x.valueTwo + 1 - G_INPUTS.grid_x.valueOne) *
            (G_INPUTS.grid_y.valueTwo + 1 - G_INPUTS.grid_y.valueOne) *
            math.ceil((G_INPUTS.grid_z.valueTwo + 1 - G_INPUTS.grid_z.valueOne) / G_INPUTS.grid_z_step.valueOne) *
            len(G_INPUTS.wall_thickness) * amount_divisions)
        
    def generate_gif(self):        
        self.__progress_dialog.reset()
        to_generate = len(self.__screenshot_filenames) if self.__generate_gif_all else 0
        to_generate += sum(map(lambda _: len(_), self.__screenshot_z_filenames)) if self.__generate_gif_row else 0
        
        todo = (1 if self.__generate_gif_all else 0) + (len(self.__screenshot_z_filenames) if self.__generate_gif_row else 0)
        current = 0
        tpl_msg = 'Generated {current} / {todo} GIFs. Read images: %v / %m (%p%)'
        self.__progress_dialog.show("Generating GIFs... (this will take a while!)", tpl_msg.format(current=current, todo=todo), 0, to_generate, 1)
        adsk.doEvents()
    
        max_frames_per_gif = G_INPUTS.max_frames_per_gif.value
        gif_fps = G_INPUTS.gif_fps.value
        gif_lossy = G_INPUTS.gif_lossy.value
        gif_optimize = G_INPUTS.gif_optimize.value
        gif_colors = G_INPUTS.gif_colors.value

        gif_folder=f"{self.__export_folder}/gif"
        os.makedirs(gif_folder, exist_ok=True)
        if self.__generate_gif_all:
            if self.create_export_gif(self.__screenshot_filenames, f"{gif_folder}/complete-{datetime.now().strftime("%Y-%m-%dT%H-%M-%S")}.gif",
                max_frames_per_gif, gif_fps, gif_optimize, gif_lossy, gif_colors):
                current += 1
                self.__progress_dialog.message = tpl_msg.format(current=current, todo=todo)
            adsk.doEvents()

        if self.__generate_gif_row:
            for zi, zlist in enumerate(self.__screenshot_z_filenames):
                if self.was_cancelled():
                    break

                if self.create_export_gif(zlist, f"{gif_folder}/z{self.calc_z(zi):02}-{datetime.now().strftime("%Y-%m-%dT%H-%M-%S")}.gif",
                    max_frames_per_gif, gif_fps, gif_optimize, gif_lossy, gif_colors):
                    current += 1
                    self.__progress_dialog.message = tpl_msg.format(current=current, todo=todo)
                adsk.doEvents()

    def generate_zip(self):
        if not G_INPUTS.zip.value:
            return
        
        zip_folder=f"{self.__export_folder}/zip"
        os.makedirs(zip_folder, exist_ok=True)

        self.__progress_dialog.reset()
        todo = len(self.__list_ww) * len(self.__range_div) * len(self.__range_z)
        current = 0
        tpl_msg = 'Generated {current} / {todo} ZIPs. Processed files: %v / %m (%p%)'
        self.__progress_dialog.show("Generating ZIP... (this will take a while!)", tpl_msg.format(current=current, todo=todo), 0, self.get_total_processed_stl(), 1)
        adsk.doEvents()

        for wall_width in self.__list_ww:
            for divisions in self.__range_div:

                if COPY_UPLOAD_WORTHY_STLS:
                    self.copy_upload_worthy_stls(wall_width, divisions)
                    continue # no zip in case of copy!

                for z in self.__range_z:
                    zip_variant_folder = TPL_VARIANT_FOLDER.format(folder=self.__export_folder, wall_width=wall_width, divisions=divisions)
                    zip_destination = f"{zip_folder}/Gridfinity_Bin1.2_Z{z:02}WW{wall_width}_D{divisions:02}.zip"
                    if self.zip_stl_files(zip_variant_folder, z, zip_destination):
                        current += 1
                        self.__progress_dialog.message = tpl_msg.format(current=current, todo=todo)
                    adsk.doEvents()

    def is_useless_bin(self, x: int, divisions: int):
        return ((x == 1 and divisions > 2)
            or (x < 2 and divisions > 4)
            or (x < 3 and divisions > 5)
            or (x < 4 and divisions > 6)
            or (x < 5 and divisions > 8)
            or (x < 7 and divisions > 9)
            or (x < 10 and divisions > 10))

    def do_export(self):
        if self.is_exporting():
            self.stop_exporting()
            return
        self.__exporting = True
        
        if not self.setup_export_folder():
            self.stop_exporting()
            return

        if self.__progress_dialog == None:
            self.__progress_dialog = G_UI.createProgressDialog()
            self.__progress_dialog.cancelButtonText = 'Abort'
            self.__progress_dialog.isBackgroundTranslucent = False
            self.__progress_dialog.isCancelButtonShown = True
        
        self.__progress_dialog.show('Exporting', 'Exported %v / ~%m (%p%)', 0, 100, 1)
        self.__progress_dialog.reset()
        
        # Get the root component of the active design
        self.__design = adsk.fusion.Design.cast(G_APP.activeProduct)

        # Parameters
        self.setup_fusion_params(self.__design)
        self.setup_ui_params()
        
        self.__amount = 0
        self.__skipped = 0
        self.__screenshot_filenames.clear()
        self.__screenshot_z_filenames.clear()

        # parameter list
        param_x = self.__design.allParameters.itemByName(PARAMS.X)
        param_y = self.__design.allParameters.itemByName(PARAMS.Y)
        param_z = self.__design.allParameters.itemByName(PARAMS.Z)

        param_wall = self.__design.allParameters.itemByName(PARAMS.WALL_THICKNESS)
        param_divisions = self.__design.allParameters.itemByName(PARAMS.DIVISIONS)

        self.__bin_parameter_list = [param_x, param_y, param_z, param_wall, param_divisions]
        time_start = timer()
   
        try:
            self.__export_manager = adsk.fusion.ExportManager.cast(self.__design.exportManager)
            self.__do_export_loop()
        except KeyboardInterrupt:
            self.stop_exporting()
            G_UI.messageBox(f"Aborted and created {self.__amount} stl files")
            return
        except:
            self.stop_exporting()
            G_UI.messageBox('Export Error:\n{}'.format(traceback.format_exc()))
            return
        finally:
            self.__progress_dialog.progressValue = self.__progress_dialog.maximumValue


        time_end = timer()
        time_delta = timedelta(seconds=time_end - time_start)

        res_msgbox = G_UI.messageBox(f"Finished and created {self.__amount} stl files. Export took {time_delta}. Continue with GIF / ZIP (if checked) after ok...")
        if res_msgbox == adsk.core.DialogResults.DialogOK or res_msgbox == adsk.core.DialogResults.DialogYes:
            if self.__generate_gif_all or self.__generate_gif_row:
                self.generate_gif()
            self.generate_zip() 

        self.stop_exporting()
        res_msgbox = G_UI.messageBox("Everything is done 🥳. Open export directory now?", "Done ✅", adsk.core.MessageBoxButtonTypes.YesNoButtonType); 
        if res_msgbox == adsk.core.DialogResults.DialogYes:
            self.view_dir_in_explorer(self.__export_folder)

    def __do_export_loop(self):
        for x in self.__range_x:
            for y in self.__range_y:
                for zi, z in enumerate(self.__range_z):
                    if len(self.__screenshot_z_filenames) == zi:
                        self.__screenshot_z_filenames.append([])

                    for wi, wall_width in enumerate(self.__list_ww):
                        for di, divisions in enumerate(self.__range_div):
                            if not self.is_exporting() or self.was_cancelled():
                                raise KeyboardInterrupt

                            self.__do_export_loop_step(x, y, z, zi, wall_width, wi, divisions)

    def __do_export_loop_step_params(self, x: int, y: int, z: int, wall_width: float, divisions: int):
        self.__design.modifyParameters(self.__bin_parameter_list, [
            adsk.core.ValueInput.createByReal(x),
            adsk.core.ValueInput.createByReal(y),
            adsk.core.ValueInput.createByReal(z),
            adsk.core.ValueInput.createByString(str(f"{wall_width} mm")),
            adsk.core.ValueInput.createByReal(divisions)
        ])

        # Process events (twice to be sure) so the file is up-2-date
        # G_APP.fireCustomEvent('thomasa88_ParametricText_Ext_Update')
        adsk.doEvents()
        adsk.doEvents()

    def __do_export_loop_step(self, x: int, y: int, z: int, z_index: int, wall_width: float, wall_index: int, divisions: int):
        variant_folder = TPL_VARIANT_FOLDER.format(folder=self.__export_folder, wall_width=wall_width, divisions=divisions)
        variant_name = TPL_VARIANT_NAME.format(x=f"{x:02}", y=f"{y:02}", z=f"{z:02}", wall_width=wall_width, divisions=f"{divisions:02}")
        stl_filename = f"{variant_folder}/{variant_name}.stl"

        if self.__generate_no_useless and self.is_useless_bin(x, divisions):
            self.__skipped += 1
            return

        filename_screenshot = f"{variant_name}.jpg"
        fullpath_screenshot = f"{self.get_screenshot_folder()}/{filename_screenshot}"
        
        should_skip_stl = self.__skip_existing_stl and os.path.isfile(stl_filename)
        should_generate_screenshot = not self.__skip_image_creation and wall_index == 0 # only for first wall width
        screenshot_exists_already = should_generate_screenshot and should_skip_stl and os.path.isfile(fullpath_screenshot)

        require_parameter_change = not should_skip_stl or (not screenshot_exists_already and should_generate_screenshot)
        if require_parameter_change:
            self.__do_export_loop_step_params(x, y, z, wall_width, divisions)
        
        if not should_skip_stl:
            os.makedirs(variant_folder, exist_ok=True)
            
            root_comp = self.__design.rootComponent
            stl_ops = self.__export_manager.createSTLExportOptions(root_comp, stl_filename)
            stl_ops.meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementMedium

            # only move camera if we have to
            if should_generate_screenshot and not screenshot_exists_already:
                G_APP.activeViewport.setCurrentAsHome(True)
                G_APP.activeViewport.goHome(False)

            self.__export_manager.execute(stl_ops)
            self.__amount += 1
        else:
            self.__skipped += 1

        if should_generate_screenshot and (screenshot_exists_already or G_APP.activeViewport.saveAsImageFile(fullpath_screenshot, INIT_SCREENSHOT_W, INIT_SCREENSHOT_H)):   
            if self.__generate_gif_all:
                self.__screenshot_filenames.append(filename_screenshot)
            if self.__generate_gif_row:
                self.__screenshot_z_filenames[z_index].append(filename_screenshot)
        
        self.__progress_dialog.progressValue = self.get_total_processed_stl()
        adsk.doEvents()
        print(f"processed: {stl_filename}")

    def __cm_into_mm(self, val: float):
        return round(val * 10, 2)

    def __read_gif_images(self, files: List[str]):
        images = []
        for i, name in enumerate(files):
            images.append(imageio.v3.imread(f"{self.get_screenshot_folder()}/{name}"))
            self.__progress_dialog.progressValue += 1
            if i % 3 == 0:
                adsk.doEvents()
                if self.was_cancelled():
                    raise KeyboardInterrupt
        return images

    # https://www.lcdf.org/gifsicle/man.html
    def create_export_gif(self, filenames: List[str], out_file_base: str, max_frames: int, fps: int = 6, optimize: int = 3, lossy: int = 80, colors: int = 256) -> bool:
        splitted = list(itertools.batched(filenames, max_frames)) if max_frames > 0 else [filenames]

        try:
            for fi, files in enumerate(splitted):
                if self.was_cancelled():
                    raise KeyboardInterrupt

                images = self.__read_gif_images(files)                
                out_file = out_file_base if fi == 0 else f"{out_file_base.removesuffix('.gif')}-part{fi+1}.gif"

                imageio.v3.imwrite(out_file, images, fps=fps) # duration seems to be broken, using fps
                pygifsicle.gifsicle(out_file, optimize=False, colors=colors, options=['--loop', f'--lossy={lossy}', f'--optimize={optimize}'])
                print(f"generated gif: {out_file}")

        except KeyboardInterrupt:
            return False
        except:
            G_UI.messageBox('GIF creation error:\n{}'.format(traceback.format_exc()))
            return False
        return True

    def copy_upload_worthy_stls(self, wall_width: int, divisions: int):
        variant_folder = TPL_VARIANT_FOLDER.format(folder=self.__export_folder, wall_width=wall_width, divisions=divisions)
        upload_folder = TPL_VARIANT_FOLDER.format(folder=f"{self.__export_folder}/todo-upload/", wall_width=wall_width, divisions=divisions)
        os.makedirs(upload_folder, exist_ok=True)

        re_search = re.compile(r"_(0[1-6])x(0[1-6])x(\d+)_")
        files = [f for f in os.listdir(variant_folder) if re_search.search(f)]
        for filename in files:
            shutil.copyfile(f"{variant_folder}/{filename}", f"{upload_folder}/{filename}")

    def zip_stl_files(self, base: str, z: int, destination: str) -> bool:
        variant_name = TPL_VARIANT_NAME.format(x='*', y='*', z=f"{z:02}", wall_width='*', divisions='*')
        files = glob.glob(f"{base}/{variant_name}.stl", recursive=True, include_hidden=True)

        zf = zipfile.ZipFile(destination, mode="w")
        try:
            for i, path in enumerate(files):
                zf.write(path, os.path.basename(path), compress_type=zipfile.ZIP_DEFLATED)
                self.__progress_dialog.progressValue += 1
                adsk.doEvents()
                if self.was_cancelled():
                    raise KeyboardInterrupt
        except KeyboardInterrupt:
            return False
        except FileNotFoundError:
            G_UI.messageBox('Error during zip:\n{}'.format(traceback.format_exc()))
            return False
        finally:
            zf.close()
        print(f"created zip: {destination}")
        return True
    
    def view_dir_in_explorer(self, path):
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", path]) # mac
        else:
            subprocess.run(["xdg-open", path]) # Linux

G_EXPORTER: GridfinityBinExporter | None = None

def update_sliders(slider_inputs: adsk.core.CommandInputs, control_input_id: str, type: Literal['wall', 'division']):
    """
    Add / remove sliders from group
    """
    spinner: adsk.core.IntegerSliderCommandInput = slider_inputs.itemById(control_input_id)
    value = spinner.valueOne
    # check ranges
    if value > spinner.maximumValue or value < spinner.minimumValue:
        return
    
    match type:
        case 'wall':
            G_INPUTS.clear_wall_thickness()

    # delete all sliders we have
    to_remove: List[adsk.core.CommandInput] = []
    for i in range(slider_inputs.count):
        j = slider_inputs.item(i)
        match type:
            case 'wall' if j.objectType == adsk.core.FloatSpinnerCommandInput.classType():
                to_remove.append(j)

    for j in to_remove:
        j.deleteMe()

    # create new ones with range depending on total number
    for i in range(1, value+1):
        cid = str(i)
        match type:
            case 'wall':
                thickness = 0.9
                if i <= 2:
                    thickness = 1.5 if i == 1 else 1.2
                G_INPUTS.wall_thickness.append(
                    slider_inputs.addFloatSpinnerCommandInput(f"wall-thickness-{cid}", f"Wall Thickness #{cid}", 'mm', 0.4, 1.75, 1, thickness)
                )
         
class GridfinityBinExporterCommandDestroyHandler(adsk.core.CommandEventHandler):
    """
    Event handler that reacts to when the command is destroyed. This terminates the script.  
    """
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            global G_EXPORTER
            if G_EXPORTER != None:
                G_EXPORTER.stop_exporting()
                G_EXPORTER = None

            # When the command is done, terminate the script
            # This will release all globals which will remove all event handlers
            adsk.terminate()
        except:
            G_UI.messageBox('Failed to Destroy:\n{}'.format(traceback.format_exc()))

class GridfinityBinExporterCommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    """
    Event handler that reacts to any changes the user makes to any of the command inputs.
    """
    def __init__(self):
        super().__init__()

    def setup(self):
        self.handleImageGifCheckboxes()

    def handleImageGifCheckboxes(self):
        gif_all = G_INPUTS.cbox_gif_all.value
        gif_z = G_INPUTS.cbox_gif_z.value
        create_imgs = G_INPUTS.cbox_create_images.value

        G_INPUTS.cbox_gif_all.isEnabled = create_imgs
        G_INPUTS.cbox_gif_z.isEnabled = create_imgs
        G_INPUTS.cbox_create_images.isEnabled = not gif_z and not gif_all

    def notify(self, args):
        global G_EXPORTER

        try:
            event_args = adsk.core.InputChangedEventArgs.cast(args)
            event_input = event_args.input
            match event_input.id:
                case IDS.CBOX_CREATE_IMAGE | IDS.CBOX_GIF_ALL | IDS.CBOX_GIF_Z:
                    self.handleImageGifCheckboxes()
                case IDS.SLIDER_WALL:
                    grp = adsk.core.GroupCommandInput.cast(event_input.parentCommandInput)
                    inputs = grp.children
                    update_sliders(inputs, event_input.id, 'wall')
                case IDS.BTN_EXPORT:
                    if G_EXPORTER:
                        G_EXPORTER.stop_exporting()
                        G_EXPORTER = None
                        return
  
                    G_EXPORTER = GridfinityBinExporter()
                    G_EXPORTER.do_export()
                    G_EXPORTER = None
        except:
            _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

class GridfinityBinExporterCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    """
    Setup UI
    """
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            # Get the command that was created.
            cmd = adsk.core.Command.cast(args.command)
            cmd.setDialogMinimumSize(360, 200)
            cmd.setDialogInitialSize(500, 500)

            # Connect to the command destroyed event.
            on_destroy = GridfinityBinExporterCommandDestroyHandler()
            cmd.destroy.add(on_destroy)
            G_HANDLERS.append(on_destroy)

            # Connect to the input changed event.           
            on_input_changed = GridfinityBinExporterCommandInputChangedHandler()
            cmd.inputChanged.add(on_input_changed)
            G_HANDLERS.append(on_input_changed)    

            # Get the CommandInputs collection associated with the command.
            inputs = cmd.commandInputs

            # Create a tab input.
            tab1 = inputs.addTabCommandInput('tab_1', 'Exporter')
            tab1_childs = tab1.children

            min_grid = 1
            max_grid = 20
            G_INPUTS.grid_x = tab1_childs.addIntegerSliderCommandInput('grid-x', 'X', min_grid, max_grid, True)
            G_INPUTS.grid_x.valueTwo = INIT_X_END
            G_INPUTS.grid_y = tab1_childs.addIntegerSliderCommandInput('grid-y', 'Y', min_grid, max_grid, True)
            G_INPUTS.grid_y.valueTwo = INIT_Y_END
            G_INPUTS.grid_z = tab1_childs.addIntegerSliderCommandInput('grid-z', 'Z', 3, 21, True)
            G_INPUTS.grid_z.valueTwo = INIT_Z_END

            G_INPUTS.grid_z_step = tab1_childs.addIntegerSliderCommandInput('grid-z-step', 'Z Step', 1, 5)
            G_INPUTS.grid_z_step.valueOne = INIT_Z_STEPS
        
            # Wall Thickness
            slider_group = tab1_childs.addGroupCommandInput('group_wall_thickness', "Wall Thickness")
            slider_inputs = slider_group.children
            slider = slider_inputs.addIntegerSliderCommandInput(IDS.SLIDER_WALL, "Wall Thickness Variations", 1, 10)
            slider.valueOne = 3
            update_sliders(slider_inputs, IDS.SLIDER_WALL, 'wall')

            G_INPUTS.division = tab1_childs.addIntegerSliderCommandInput(IDS.SLIDER_DIVISION, "Divisions", 1, 15, True)
            G_INPUTS.division.valueTwo = 6
 
            # Scoop
            scoops_group = tab1_childs.addGroupCommandInput('group_scoops', "Scoop")
            scoops_group_inputs = scoops_group.children
            G_INPUTS.scoop_radius = scoops_group_inputs.addIntegerSliderCommandInput('slider-scoop-radius', "Scoop Curve (in mm)", 0, 50)
            G_INPUTS.scoop_radius.tooltip = 'Radius of the bottom (and top) scoop curve in mm'
            G_INPUTS.scoop_radius.tooltipDescription = 'Set 0 to disable it, default is 10 mm. Z <3 should be max 10, Z<4 max 15, ...'
            G_INPUTS.scoop_radius.valueOne = 10

            # Magnets
            group_magnets = tab1_childs.addGroupCommandInput('group-gif', 'Magnets')

            G_INPUTS.mag_diameter = group_magnets.children.addFloatSpinnerCommandInput(
                 'magnet-diameter', 'Magnet Diameter', 'mm', 2, 8.2, 1, 6.1)

            G_INPUTS.mag_rem_diameter  = group_magnets.children.addFloatSpinnerCommandInput(
                 'magnet-remove-diameter', 'Magnet Remove Hole Diameter', 'mm', 0, 6, 0.1, 3)

            G_INPUTS.mag_depth = group_magnets.children.addFloatSpinnerCommandInput(
                 'magnet-height', 'Magnet Depth', 'mm', 1, 4, 0.1, 2.4)

            group_gif = tab1_childs.addGroupCommandInput('group-gif', 'GIF creation')
            
            G_INPUTS.cbox_gif_all = group_gif.children.addBoolValueInput(IDS.CBOX_GIF_ALL, 'Generate Complete', True, '', False)
            G_INPUTS.cbox_gif_z = group_gif.children.addBoolValueInput(IDS.CBOX_GIF_Z, 'Generate per Z-Axis', True, '', False)
            G_INPUTS.max_frames_per_gif = group_gif.children.addIntegerSpinnerCommandInput('spin-max-gif-frames', 'Max frames per GIF', 0, 50*1000, 1, 0)

            G_INPUTS.gif_fps = group_gif.children.addIntegerSpinnerCommandInput('spin-gif-fps', 'FPS', 0, 60, 1, 6)
            G_INPUTS.gif_lossy = group_gif.children.addIntegerSpinnerCommandInput('spin-gif-fps', 'Lossy', 0, 200, 1, 100)
            G_INPUTS.gif_optimize = group_gif.children.addIntegerSpinnerCommandInput('spin-gif-fps', 'Optimize', 0, 3, 1, 3)
            G_INPUTS.gif_colors = group_gif.children.addIntegerSpinnerCommandInput('spin-gif-colors', 'Colors', 32, 256, 32, 128)
           
            # ----

            G_INPUTS.cbox_create_images = tab1_childs.addBoolValueInput(IDS.CBOX_CREATE_IMAGE, 'Create preview image', True, '', True)
            G_INPUTS.cbox_create_images.tooltip = 'Create a screenshot image for each exported bin. Also required by GIF creation.'

            G_INPUTS.cbox_useless = tab1_childs.addBoolValueInput('cbox-usless', 'Prevent useless bins', True, '', True)
            G_INPUTS.cbox_useless.tooltip = 'Do not generate bins with crazy amounts of divisions which are probably useless'
 
            G_INPUTS.skip_existing = tab1_childs.addBoolValueInput('cbox-skip-existing', 'Skip Existing STL', True, '', True)
            G_INPUTS.skip_existing.tooltip = 'Skip exporting if the STL file already exists'

            G_INPUTS.zip = tab1_childs.addBoolValueInput('cbox-zip', 'ZIP files', True, '', False)

            tab1_childs.addBoolValueInput(IDS.BTN_EXPORT, 'Export', False, '', True)

            on_input_changed.setup()
        except:
            G_UI.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def run(context):
    """
    Main starting point
    """
    try:
        global G_APP, G_UI
        G_APP = adsk.core.Application.get()
        G_UI = G_APP.userInterface

        # Get the existing command definition or create it if it doesn't already exist.
        cmd_definition = G_UI.commandDefinitions.itemById('cmd-gridfinitybin-exporter')
        if not cmd_definition:
            cmd_definition = G_UI.commandDefinitions.addButtonDefinition('cmd-gridfinitybin-exporter', 'Gridfinity Bin 1.2 Exporter', 'Export a lot of STL files :)')

        # Connect to the command created event.
        on_command_created = GridfinityBinExporterCommandCreatedHandler()
        cmd_definition.commandCreated.add(on_command_created)
        G_HANDLERS.append(on_command_created)

        # Execute the command definition.
        cmd_definition.execute()

        # Prevent this module from being terminated when the script returns, because we are waiting for event handlers to fire.
        adsk.autoTerminate(False)
    except:
        if G_UI:
            G_UI.messageBox('Failed:\n{}'.format(traceback.format_exc()))
