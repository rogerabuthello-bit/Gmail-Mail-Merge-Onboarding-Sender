import re
from typing import Any

from jinja2 import Environment, StrictUndefined


VARIABLE_PATTERN = re.compile(r"{{\s*([^{}]+?)\s*}}")


def detect_template_variables(*templates: str) -> list[str]:
    seen: dict[str, None] = {}
    for template in templates:
        if not template:
            continue
        for match in VARIABLE_PATTERN.findall(template):
            variable = match.strip()
            if variable and variable not in seen:
                seen[variable] = None
    return list(seen.keys())


def auto_map_variables(variables: list[str], columns: list[str]) -> dict[str, str | None]:
    normalized_columns = {normalize_name(column): column for column in columns}
    mapping: dict[str, str | None] = {}
    for variable in variables:
        mapping[variable] = normalized_columns.get(normalize_name(variable))
    return mapping


def normalize_name(value: str) -> str:
    return re.sub(r"\s+", " ", str(value).strip()).casefold()


def build_render_context(
    row: dict[str, Any],
    variable_mapping: dict[str, str | None],
) -> tuple[dict[str, Any], dict[str, str]]:
    context: dict[str, Any] = {}
    token_mapping: dict[str, str] = {}

    for index, (variable, column_name) in enumerate(variable_mapping.items()):
        token = f"field_{index}"
        token_mapping[variable] = token

        if column_name is None:
            value = ""
        else:
            raw_value = row.get(column_name, "")
            value = "" if raw_value is None else str(raw_value)

        context[token] = value

    return context, token_mapping


def convert_template_to_jinja(
    template: str,
    token_mapping: dict[str, str],
) -> str:
    converted = template
    for variable, token in token_mapping.items():
        pattern = re.compile(r"{{\s*" + re.escape(variable) + r"\s*}}")
        converted = pattern.sub("{{ " + token + " }}", converted)
    return converted


def render_template(
    template: str,
    row: dict[str, Any],
    variable_mapping: dict[str, str | None],
) -> str:
    context, token_mapping = build_render_context(row, variable_mapping)
    converted_template = convert_template_to_jinja(template, token_mapping)
    environment = Environment(undefined=StrictUndefined, autoescape=False)
    return environment.from_string(converted_template).render(**context)


def append_signature(html_body: str, signature_html: str, include_signature: bool) -> str:
    if not include_signature or not signature_html.strip():
        return html_body
    return f"{html_body}<br><br>{signature_html}"
