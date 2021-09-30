import argparse, json
from pathlib import Path

from .template import DocxTemplate, TemplateError

TEMPLATE_ARG = 'template_path'
JSON_ARG = 'json_path'
OUTPUT_ARG = 'output_filename'
OVERWRITE_ARG = 'overwrite'
QUIET_ARG = 'quiet'


def make_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        usage=f'python -m docxtpl [-h] [-o] [-q] {TEMPLATE_ARG} {JSON_ARG} {OUTPUT_ARG}',
        description='Make docx file from existing template docx and json data.')
    parser.add_argument(TEMPLATE_ARG,
                        type=str,
                        help='The path to the template docx file.')
    parser.add_argument(JSON_ARG,
                        type=str,
                        help='The path to the json file with the data.')
    parser.add_argument(OUTPUT_ARG,
                        type=str,
                        help='The filename to save the generated docx.')
    parser.add_argument('-' + OVERWRITE_ARG[0], '--' + OVERWRITE_ARG,
                        action='store_true',
                        help='If output file already exists, overwrites without asking for confirmation')
    parser.add_argument('-' + QUIET_ARG[0], '--' + QUIET_ARG,
                        action='store_true',
                        help='Do not display unnecessary messages')
    return parser


def get_args(parser: argparse.ArgumentParser) -> dict:
    try:
        parsed_args = vars(parser.parse_args())
        return parsed_args
    # Argument errors raise a SystemExit with code 2. Normal usage of the
    # --help or -h flag raises a SystemExit with code 0.
    except SystemExit as e:
        if e.code == 0:
            raise SystemExit from e
        else:
            raise RuntimeError(f'Correct usage is:\n{parser.usage}') from e


def is_argument_valid(arg_name: str, arg_value: str,overwrite: bool) -> bool:
    # Basic checks for the arguments
    if arg_name == TEMPLATE_ARG:
        return Path(arg_value).is_file() and arg_value.endswith('.docx')
    elif arg_name == JSON_ARG:
        return Path(arg_value).is_file() and arg_value.endswith('.json')
    elif arg_name == OUTPUT_ARG:
        return arg_value.endswith('.docx') and check_exists_ask_overwrite(
            arg_value, overwrite)
    elif arg_name in [OVERWRITE_ARG, QUIET_ARG]:
        return arg_value in [True, False]


def check_exists_ask_overwrite(arg_value:str, overwrite: bool) -> bool:
    # If output file does not exist or command was run with overwrite option,
    # returns True, else asks for overwrite confirmation. If overwrite is
    # confirmed returns True, else raises FileExistsError.
    if Path(arg_value).exists() and not overwrite:
        try:
            if input(f'File {arg_value} already exists, would you like to overwrite the existing file? (y/n)').lower() == 'y':
                return True
            else:
                raise FileExistsError
        except FileExistsError as e:
            raise RuntimeError(f'File {arg_value} already exists, please choose a different name.') from e
    else:
        return True


def validate_all_args(parsed_args:dict) -> None:
    overwrite = parsed_args[OVERWRITE_ARG]
    # Raises AssertionError if any of the arguments is not validated
    try:
        for arg_name, arg_value in parsed_args.items():
            if not is_argument_valid(arg_name, arg_value,overwrite):
                raise AssertionError
    except AssertionError as e:
        raise RuntimeError(
            f'The specified {arg_name} "{arg_value}" is not valid.') from e


def get_json_data(json_path: Path) -> dict:
    with open(json_path) as file:
        try:
            json_data = json.load(file)
            return json_data
        except json.JSONDecodeError as e:
            print(
                f'There was an error on line {e.lineno}, column {e.colno} while trying to parse file {json_path}')
            raise RuntimeError('Failed to get json data.') from e


def make_docxtemplate(template_path: Path) -> DocxTemplate:
    try:
        return DocxTemplate(template_path)
    except TemplateError as e:
        raise RuntimeError('Could not create docx template.') from e


def render_docx(doc:DocxTemplate, json_data: dict) -> DocxTemplate:
    try:
        doc.render(json_data)
        return doc
    except TemplateError as e:
        raise RuntimeError(f'An error ocurred while trying to render the docx') from e


def save_file(doc: DocxTemplate, parsed_args: dict) -> None:
    try:
        output_path = parsed_args[OUTPUT_ARG]
        doc.save(output_path)
        if not parsed_args[QUIET_ARG]:
            print(f'Document successfully generated and saved at {output_path}')
    except PermissionError as e:
        print(f'{e.strerror}. Could not save file {e.filename}.')
        raise RuntimeError('Failed to save file.') from e


def main() -> None:
    parser = make_arg_parser()
    # Everything is in a try-except block that cacthes a RuntimeError that is
    # raised if any of the individual functions called cause an error
    # themselves, terminating the main function.
    parsed_args = get_args(parser)
    try:
        validate_all_args(parsed_args)
        json_data = get_json_data(Path(parsed_args[JSON_ARG]).resolve())
        doc = make_docxtemplate(Path(parsed_args[TEMPLATE_ARG]).resolve())
        doc = render_docx(doc,json_data)
        save_file(doc, parsed_args)
    except RuntimeError as e:
        print('Error: '+e.__str__())
        return
    finally:
        if not parsed_args[QUIET_ARG]:
            print('Exiting program!')


if __name__ == '__main__':
    main()
