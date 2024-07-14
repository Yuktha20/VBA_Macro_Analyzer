import os
import codecs
import oletools.olevba
import re
from graphviz import Digraph

# Strings produced by olevba that mark the beginning & end of a VBA stream
STREAM_START = r'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'
STREAM_END = r'-------------------------------------------------------------------------------'

def collect_files(directory):
    """
    Collect all Excel files in the specified directory.

    Args:
        directory (str): The directory to search for Excel files.

    Returns:
        list: A list of paths to Excel files.
    """
    excel_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith((".xls", ".xlsm")):
                excel_files.append(os.path.join(root, file))
    return excel_files

def extract_macros(file_path):
    """
    Extract VBA macros from an Excel file using olevba.

    Args:
        file_path (str): The path to the Excel file.

    Returns:
        list: A list of VBA code strings.
    """
    vba_parser = oletools.olevba.VBA_Parser(file_path)
    macros = []
    if vba_parser.detect_vba_macros():
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_all_macros():
            macros.append(vba_code)
    vba_parser.close()
    return macros

def analyze_vba_code(vba_code):
    """
    Analyze the structure and logic of VBA code.

    Args:
        vba_code (str): The VBA code to analyze.

    Returns:
        dict: A dictionary with analysis results.
    """
    analysis = {
        "functions": [],
        "subroutines": [],
        "variables": [],
        "logic_flow": [],
        "data_flow": []
    }
    lines = vba_code.splitlines()

    for line in lines:
        line = line.strip()
        if line.startswith("Sub ") or line.startswith("Function "):
            analysis["logic_flow"].append(line)
        elif line.startswith("Dim "):
            analysis["variables"].append(line)
        elif re.match(r'^\w+\s*=', line):
            analysis["data_flow"].append(line)
        else:
            analysis["logic_flow"].append(line)

    return analysis

def generate_documentation(analysis, output_path):
    """
    Generate documentation from the VBA code analysis.

    Args:
        analysis (dict): The analysis results.
        output_path (str): The path to save the documentation.
    """
    with codecs.open(output_path, 'w', encoding='utf-8') as doc:
        doc.write("VBA Macro Analysis\n")
        doc.write("==================\n\n")

        doc.write("Functions and Subroutines:\n")
        for item in analysis["functions"]:
            doc.write(f"- {item}\n")
        doc.write("\n")

        doc.write("Variables:\n")
        for item in analysis["variables"]:
            doc.write(f"- {item}\n")
        doc.write("\n")

        doc.write("Logic Flow:\n")
        for item in analysis["logic_flow"]:
            doc.write(f"- {item}\n")
        doc.write("\n")

        doc.write("Data Flow:\n")
        for item in analysis["data_flow"]:
            doc.write(f"- {item}\n")

def generate_flowchart(analysis, output_path):
    """
    Generate a flowchart from the VBA code analysis using graphviz.

    Args:
        analysis (dict): The analysis results.
        output_path (str): The path to save the flowchart image.
    """
    dot = Digraph(comment='VBA Macro Flowchart')
    dot.attr('node', shape='box')

    for func in analysis['functions']:
        dot.node(func, func)

    prev_node = None
    for statement in analysis['logic_flow']:
        dot.node(statement, statement)
        if prev_node:
            dot.edge(prev_node, statement)
        prev_node = statement

    dot.render(output_path, format='png')
