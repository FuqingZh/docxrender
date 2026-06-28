from __future__ import annotations

from collections.abc import Mapping
from typing import Any, cast

from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage

from docxrender.contracts import (
    DocxTemplateContextPolicy,
    DocxTemplateImageSpec,
    DocxTemplateRenderOptions,
    DocxTemplateRenderResult,
)


def write_docx_template(
    options: DocxTemplateRenderOptions,
) -> DocxTemplateRenderResult:
    """Write a DOCX by rendering a `docxtpl` template with caller context.

    Args:
        options (DocxTemplateRenderOptions): Template rendering options.

    Returns:
        DocxTemplateRenderResult: Result containing the rendered DOCX path.

    Raises:
        FileNotFoundError: The template path does not exist.
        RuntimeError: The rendered DOCX cannot be written.
    """

    options.file_out_docx.parent.mkdir(parents=True, exist_ok=True)
    template = cast(Any, DocxTemplate(str(options.file_template)))
    context = resolve_template_context(
        context=options.context,
        context_defaults=options.context_defaults,
        context_policy=options.context_policy,
    )
    context_render = materialize_template_context(
        template=template,
        context=context,
        inline_images=options.inline_images,
    )
    template.render(context_render)
    template.save(str(options.file_out_docx))
    return DocxTemplateRenderResult(file_docx=options.file_out_docx)


def resolve_template_context(
    *,
    context: Mapping[str, Any],
    context_defaults: Mapping[str, Any],
    context_policy: DocxTemplateContextPolicy,
) -> dict[str, Any]:
    """Resolve final `docxtpl` context from caller values, defaults, and policy.

    Args:
        context (dict[str, Any] | Any): Caller-provided context mapping.
        context_defaults (dict[str, Any] | Any): Default context mapping.
        context_policy (DocxTemplateContextPolicy): Merge and validation policy.

    Returns:
        dict[str, Any]: Final merged render context.

    Raises:
        ValueError: The merge rule, conflict rule, or required keys are invalid.
    """

    if context_policy.rule_merge != "merge":
        raise ValueError(
            f"Unsupported template context merge rule: {context_policy.rule_merge!r}"
        )

    caller_context = dict(context)
    default_context = dict(context_defaults)
    if context_policy.rule_conflict == "caller_wins":
        merged_context = dict(default_context)
        merged_context.update(caller_context)
    elif context_policy.rule_conflict == "defaults_win":
        merged_context = dict(caller_context)
        merged_context.update(default_context)
    else:
        raise ValueError(
            "Unsupported template context conflict rule: "
            f"{context_policy.rule_conflict!r}"
        )

    missing_keys = tuple(
        key for key in context_policy.required_keys if key not in merged_context
    )
    if missing_keys:
        raise ValueError(
            "Missing required template context keys after merge: "
            f"required_keys={context_policy.required_keys!r} "
            f"missing_keys={missing_keys!r}"
        )
    return merged_context


def materialize_template_context(
    *,
    template: DocxTemplate,
    context: Mapping[str, Any],
    inline_images: Mapping[str, DocxTemplateImageSpec],
) -> dict[str, Any]:
    """Materialize template runtime values on the active `DocxTemplate` instance.

    Args:
        template (DocxTemplate): Active template instance for this render.
        context (Mapping[str, Any]): Resolved render context before runtime values.
        inline_images (Mapping[str, DocxTemplateImageSpec]): Image specs to bind to
            this template instance.

    Returns:
        dict[str, Any]: Runtime render context including bound image objects.

    Raises:
        ValueError: Inline-image keys conflict with ordinary context keys.
    """

    overlapping_keys = tuple(key for key in inline_images if key in context)
    if overlapping_keys:
        raise ValueError(
            "Template inline-image keys conflict with ordinary context keys: "
            f"overlapping_keys={overlapping_keys!r}"
        )
    context_runtime = dict(context)
    for key, spec in inline_images.items():
        context_runtime[key] = create_template_inline_image(template, spec)
    return context_runtime


def create_template_inline_image(
    template: DocxTemplate,
    spec: DocxTemplateImageSpec,
) -> InlineImage:
    """Create an `InlineImage` bound to the current template instance."""

    width = None if spec.width_mm is None else Mm(spec.width_mm)
    height = None if spec.height_mm is None else Mm(spec.height_mm)
    return InlineImage(
        template,
        str(spec.file_image),
        width=width,
        height=height,
        anchor=spec.anchor,
    )
