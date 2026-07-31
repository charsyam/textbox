import unicodedata


def display_width(text):
    return sum(
        0
        if unicodedata.combining(char)
        else 2
        if unicodedata.east_asian_width(char) in ("W", "F")
        else 1
        for char in str(text)
    )


def wrap_display(text, width):
    result = []
    for logical_line in str(text).splitlines() or [""]:
        current = []
        current_width = 0
        for char in logical_line:
            char_width = display_width(char)
            if current and current_width + char_width > width:
                result.append("".join(current))
                current = []
                current_width = 0
            current.append(char)
            current_width += char_width
        result.append("".join(current))
    return result or [""]


def allocate_widths(weights, natural_widths, max_total=100):
    count = len(natural_widths)
    if not count:
        return []
    weights = list(weights or [1] * count)
    if len(weights) != count or not any(weights):
        weights = [1] * count
    weights = [max(float(value or 0), 0.0) for value in weights]
    fallback = sum(value for value in weights if value) / max(
        sum(1 for value in weights if value), 1
    )
    weights = [value or fallback or 1 for value in weights]

    target = max(count * 3, min(max_total, sum(min(value, 24) for value in natural_widths)))
    widths = [3] * count
    remaining = max(target - sum(widths), 0)
    total = sum(weights)
    exact = [remaining * value / total for value in weights]
    additions = [int(value) for value in exact]
    for index, addition in enumerate(additions):
        widths[index] += addition
    leftover = remaining - sum(additions)
    order = sorted(
        range(count), key=lambda index: exact[index] - additions[index], reverse=True
    )
    for index in order[:leftover]:
        widths[index] += 1
    for index, natural in enumerate(natural_widths):
        widths[index] = min(max(widths[index], min(natural, 12)), 30)
    return widths


def render_grid(grid, title=None, column_weights=None, max_width=100):
    if not grid:
        return title or ""
    cols = max((len(row) for row in grid), default=0)
    if not cols:
        return title or ""
    normalized = [
        [str(value) if value is not None else "" for value in row]
        + [""] * (cols - len(row))
        for row in grid
    ]
    natural = [
        max([display_width(row[col]) for row in normalized] + [3])
        for col in range(cols)
    ]
    widths = allocate_widths(column_weights, natural, max_total=max_width)

    def border(left, joint, right):
        return left + joint.join("─" * (width + 2) for width in widths) + right

    lines = []
    if title:
        lines.append(title)
    lines.append(border("┌", "┬", "┐"))
    for row_index, row in enumerate(normalized):
        wrapped = [wrap_display(row[col], widths[col]) for col in range(cols)]
        height = max(len(value) for value in wrapped)
        for line_index in range(height):
            cells = []
            for col, values in enumerate(wrapped):
                value = values[line_index] if line_index < len(values) else ""
                cells.append(
                    " "
                    + value
                    + " " * (widths[col] - display_width(value))
                    + " "
                )
            lines.append("│" + "│".join(cells) + "│")
        if row_index + 1 < len(normalized):
            lines.append(border("├", "┼", "┤"))
    lines.append(border("└", "┴", "┘"))
    return "\n".join(lines)
