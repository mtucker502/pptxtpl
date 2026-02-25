"""InlineImage support for pptxtpl.

Allows replacing placeholder text in a template with an image.
The image is added to the slide's relationship parts and a reference
marker is returned for post-processing.
"""

import os


class InlineImage:
    """Represents an image to be inserted into a PowerPoint template.

    Usage::

        img = InlineImage(tpl, "photo.png", width=Inches(2))
        tpl.render({"photo": img})

    In the template, use ``{{ photo }}`` in a picture shape's alt-text or
    as a text placeholder that will be replaced with the image.

    Note: For the initial version, InlineImage stores the image path and
    dimensions for use in custom post-processing. Full automatic picture
    replacement requires manipulating blipFill references, which depends
    on the specific template structure.
    """

    def __init__(self, tpl, image_path: str, width=None, height=None):
        """
        Args:
            tpl: The PptxTemplate instance (used for accessing slide parts).
            image_path: Path to the image file.
            width: Optional width (in EMUs or pptx.util units).
            height: Optional height (in EMUs or pptx.util units).
        """
        self._tpl = tpl
        self._image_path = os.path.abspath(image_path)
        self._width = width
        self._height = height

    @property
    def image_path(self) -> str:
        return self._image_path

    @property
    def width(self):
        return self._width

    @property
    def height(self):
        return self._height

    def __str__(self) -> str:
        # Return a marker that can be identified in post-processing
        return f"[InlineImage:{self._image_path}]"
