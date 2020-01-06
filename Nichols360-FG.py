# Importing sample Fusion Command
# Could import multiple Command definitions here
# from .SoftJawStockSolid import SoftJawStockSolid
from .NewFG import FrameGenerator
# from .DemoPaletteCommand import DemoPaletteShowCommand, DemoPaletteSendCommand
# from .F360Web import F360Web

commands = []
command_definitions = []

# Define parameters for 1st command
# cmd = {
#     'cmd_name': 'SoftJaw Stock Solid',
#     'cmd_description': 'command to create soft jaws and stock solids',
#     'cmd_id': 'cmdID_demo50',
#     'cmd_resources': './resources/SoftJaw',
#     'workspace': 'FusionSolidEnvironment',
#     'toolbar_panel_id': 'Nichols 360',
#     'command_visible': True,
#     'command_promoted': True,
#     'class': SoftJawStockSolid
# }
# command_definitions.append(cmd)

# Define parameters for 2nd command
cmd = {
    'cmd_name': 'Frame Generator',
    'cmd_description': 'Frame Generator to create Structural Steel Members',
    'cmd_id': 'cmdID_demo20',
    'cmd_resources': './resources/FG',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Frame Generator',
    'command_visible': True,
    'command_promoted': True,
    'class': FrameGenerator
}
command_definitions.append(cmd)

# Define parameters for 2nd command
# cmd = {
#     'cmd_name': 'Tutorials and Help Guides',
#     'cmd_description': 'Tutorials and Help Guides for learning F360',
#     'cmd_id': 'cmdID_palette_demo2',
#     'cmd_resources': './resources/Tutorials',
#     'workspace': 'FusionSolidEnvironment',
#     'toolbar_panel_id': 'Nichols 360',
#     'command_visible': True,
#     'command_promoted': True,
#     'palette_id': 'demo_palette_id1',
#     'palette_name': 'Tutorials and Help Guides',
#     'palette_html_file_url': 'Tutorials.html',
#     'palette_is_visible': True,
#     'palette_show_close_button': True,
#     'palette_is_resizable': True,
#     'palette_width': 400,
#     'palette_height': 600,
#     'class': DemoPaletteShowCommand
# }
# command_definitions.append(cmd)

# Define parameters for 2nd command
# cmd = {
#     'cmd_name': 'F360 Web interface',
#     'cmd_description': 'Embedding the web interface into the design space',
#     'cmd_resources': './resources/f360Embed',
#     'workspace': 'FusionSolidEnvironment',
#     'toolbar_panel_id': 'Nichols 360',
#     'command_visible': True,
#     'command_promoted': True,
#     'palette_id': 'demo_palette_id3',
#     'palette_name': 'F360 Web interface',
#     'palette_html_file_url': 'f360.html',
#     'palette_is_visible': True,
#     'palette_show_close_button': True,
#     'palette_is_resizable': True,
#     'palette_width': 600,
#     'palette_height': 600,
#     'class': F360Web
# }
# command_definitions.append(cmd)

# # Define parameters for 2nd command
# cmd = {
#     'cmd_name': 'Fusion Palette Send Command',
#     'cmd_description': 'Send info to Fusion 360 Palette',
#     'cmd_id': 'cmdID_palette_send_demo',
#     'cmd_resources': './resources',
#     'workspace': 'FusionSolidEnvironment',
#     'toolbar_panel_id': 'SolidScriptsAddinsPanel',
#     'command_visible': True,
#     'command_promoted': False,
#     'palette_id': 'demo_palette_id',
#     'class': DemoPaletteSendCommand
# }
# command_definitions.append(cmd)

# Set to True to display various useful messages when debugging your app
debug = False

# Don't change anything below here:
for cmd_def in command_definitions:
    command = cmd_def['class'](cmd_def, debug)
    commands.append(command)


def run(context):
    for run_command in commands:
        run_command.on_run()


def stop(context):
    for stop_command in commands:
        stop_command.on_stop()
