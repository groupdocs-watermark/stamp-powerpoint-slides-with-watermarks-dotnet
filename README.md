# Stamp PowerPoint Slides with Watermarks in C#

[![Product Page](https://img.shields.io/badge/Product%20Page-2865E0?style=for-the-badge&logo=appveyor&logoColor=white)](https://products.groupdocs.com/watermark/net/)
[![Docs](https://img.shields.io/badge/Docs-2865E0?style=for-the-badge&logo=Hugo&logoColor=white)](https://docs.groupdocs.com/watermark/net/)
[![Blog](https://img.shields.io/badge/Blog-2865E0?style=for-the-badge&logo=WordPress&logoColor=white)](https://blog.groupdocs.com/categories/groupdocs.watermark-product-family/)
[![Free Support](https://img.shields.io/badge/Free%20Support-2865E0?style=for-the-badge&logo=Discourse&logoColor=white)](https://forum.groupdocs.com/c/watermark/19)
[![Temporary License](https://img.shields.io/badge/Temporary%20License-2865E0?style=for-the-badge&logo=rocket&logoColor=white)](https://purchase.groupdocs.com/temp-license/100354)

## Overview

A watermark on a PowerPoint slide is a shape, and that has consequences most
watermarking guides skip: a shape can be selected, dragged off the canvas, or
deleted by anyone who opens the deck. This repository works through the
GroupDocs.Watermark for .NET API surface that addresses exactly that - which
slides a watermark reaches, whether it can be edited once applied, and how to
put a mark somewhere a reader cannot simply click and delete.

Seven runnable methods cover the range, from a one-line text stamp across a
deck to watermarks burned into the pictures embedded inside a slide. Built on
.NET 8 with GroupDocs.Watermark 26.9.0, the examples are written for
developers who need watermarking that holds up after the file leaves their
hands.

## Technology Stack

- **Platform**: .NET 8.0
- **Product**: GroupDocs.Watermark for .NET 26.9.0
- **Language**: C# 12
- **Document format**: Office Open XML presentations (`.pptx`)

## Problem Statement

Marking a presentation is rarely "add the word CONFIDENTIAL and save". A title
slide may need a different mark from the body, a client deck needs the logo on
one slide rather than all forty, and a watermark that the recipient can remove
in two clicks does not satisfy anyone's compliance requirement.

PowerPoint makes this awkward because it has no watermark feature as such. A
watermark is a text box or picture that you position, rotate, scale, send
behind the content, and then hope nobody touches. Doing that by hand across a
deck is tedious; doing it across hundreds of decks is not realistic. And the
obvious programmatic approach - drop a shape on each slide - inherits the same
weakness: the shape is as removable as any other.

GroupDocs.Watermark handles the mechanics and, more usefully, exposes the
choices that matter: targeting, locking, and the option to watermark the
images inside a slide instead of the slide itself.

## Solution Overview

GroupDocs.Watermark treats the presentation as content parts you can address
individually. A `Watermarker` loads the file, a `TextWatermark` or
`ImageWatermark` describes the mark, and a `PresentationWatermarkSlideOptions`
decides where it lands. Key technical features exercised here:

- **Slide targeting**: `PresentationWatermarkSlideOptions.SlideIndex` applies a
  watermark to one zero-based slide; omitting the options applies it deck-wide.
- **Text and image marks**: `TextWatermark` with a `Font`, or `ImageWatermark`
  reading a logo from disk, both positioned by the same alignment properties.
- **Relative sizing**: `SizingType.ScaleToParentDimensions` with `ScaleFactor`
  scales the mark against the slide, so one value works at 4:3 and 16:9.
- **Tamper resistance**: `IsLocked` blocks editing the shape in PowerPoint and
  `ProtectWithUnreadableCharacters` defeats Find and Replace.
- **Embedded-image watermarking**: `PresentationSlide.FindImages()` returns the
  pictures on a slide so the mark can be rasterised into the image data, where
  deleting a shape will not remove it.

## Prerequisites

Before running these examples, ensure you have:

- **Runtime** - .NET 8.0 SDK or later
- **Package** - GroupDocs.Watermark 26.9.0 (restored from NuGet)
- **Licence** - a GroupDocs licence or [temporary licence](https://purchase.groupdocs.com/temp-license/100354);
  without one the library runs in evaluation mode and marks the output
- **Input files** - a `.pptx` presentation and an image for the logo example,
  both included under `documents/`

## Getting Started

### Installation

**Using Package Manager:**

```bash
dotnet add package GroupDocs.Watermark --version 26.9.0
```

**Manual Installation:**

Clone the repository and run `dotnet restore` in the project folder; the
package reference in the `.csproj` pulls everything else.

### Configuration

1. Open `Program.cs` and find the `ApplyLicense` method.
2. Replace `REPLACE_WITH_YOUR_LICENSE_PATH` with the path to your `.lic` file.
   Without a licence the program still runs - it prints a notice and produces
   evaluation-marked output.
3. Run `dotnet run`. Results are written to the `Result/` folder, one `.pptx`
   per technique.

## Repository Structure

```
stamp-powerpoint-slides-with-watermarks-dotnet/
│
├── AddWatermarkToPowerPointDemo.csproj
├── Program.cs
└── documents/
    ├── sample.pptx
    └── logo.png
```

### File Descriptions

- **AddWatermarkToPowerPointDemo.csproj** - project descriptor targeting
  `net8.0` with the single GroupDocs.Watermark package reference. The
  `documents/` folder is copied to the output directory on build.
- **Program.cs** - all seven watermarking methods plus licence setup and a
  `Main` that runs each in turn and reports the file it produced.
- **documents/sample.pptx** - a three-slide presentation. The third slide
  holds an embedded picture, which is what the `FindImages` example needs.
- **documents/logo.png** - the source image for the image-watermark example.

## Code Implementation

### Implementation: Read the slide dimensions

Positioning a watermark in absolute coordinates requires knowing the canvas
size. This reads it from the presentation rather than assuming 4:3 or 16:9.

```csharp
var loadOptions = new PresentationLoadOptions();
using var watermarker = new Watermarker(sourcePath, loadOptions);

PresentationContent content = watermarker.GetContent<PresentationContent>();
return (content.SlideWidth, content.SlideHeight);
```

#### Technical Details

`GetContent<PresentationContent>()` casts the loaded document to the
presentation-specific content model, which exposes slide-level members the
generic `Watermarker` API does not. Reading dimensions costs a load but no
save, so it is safe to call before deciding how to size a mark.

**Key Components:**

- `Watermarker`: loads a document and owns the handle until disposed
- `PresentationLoadOptions`: tells the library to treat the file as a presentation
- `PresentationContent`: presentation-specific view over the loaded document

Output: slide width and height in points, e.g. `720 x 540`.

---

### Implementation: Text watermark across the whole deck

The baseline. No slide options at all, which GroupDocs.Watermark reads as
"every slide".

```csharp
var loadOptions = new PresentationLoadOptions();
using var watermarker = new Watermarker(sourcePath, loadOptions);

var watermark = new TextWatermark(text, new Font("Arial", 42))
{
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Center,
    Opacity = 0.35,
};

watermarker.Add(watermark);
watermarker.Save(outputPath);
```

#### Technical Details

`Add` with a watermark and no options walks the deck and places the shape on
each slide. The alignment properties are relative to the slide, so the mark
stays centred whatever the aspect ratio. `Save` writes to a new path - the
source file is never modified, which matters when the same deck feeds several
techniques.

**Key Components:**

- `TextWatermark`: the mark, described by text plus a `Font`
- `Opacity`: 0.0 to 1.0; 0.35 reads as a watermark rather than a headline
- `Watermarker.Save`: writes the modified document to a new file

**Parameters:**

- `text`: the watermark string, e.g. `CONFIDENTIAL`
- `outputPath`: destination `.pptx`; the source is left untouched

Output: a copy of the deck with the mark centred on every slide.

---

### Implementation: Text watermark on one slide

Targeting a single slide, which is where most real requirements land - mark the
title slide, or the one with the pricing table.

```csharp
var loadOptions = new PresentationLoadOptions();
using var watermarker = new Watermarker(sourcePath, loadOptions);

var watermark = new TextWatermark(text, new Font("Arial", 36))
{
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Top,
};

var options = new PresentationWatermarkSlideOptions { SlideIndex = slideIndex };

watermarker.Add(watermark, options);
watermarker.Save(outputPath);
```

#### Technical Details

`SlideIndex` is zero-based: `0` is the first slide. An index beyond the deck
throws, so validate against `PresentationContent.Slides.Count` when the value
comes from configuration or user input. Everything else about the watermark is
unchanged - targeting is purely a matter of which options object you pass.

**Key Components:**

- `PresentationWatermarkSlideOptions`: the slide-scoped options object
- `SlideIndex`: zero-based index of the target slide
- `VerticalAlignment.Top`: keeps the mark clear of slide content

**Parameters:**

- `slideIndex`: which slide receives the watermark, counting from 0

Output: one slide marked, the rest of the deck byte-for-byte unchanged.

---

### Implementation: Image watermark on one slide

A logo instead of text, scaled against the slide rather than fixed in points.

```csharp
var loadOptions = new PresentationLoadOptions();
using var watermarker = new Watermarker(sourcePath, loadOptions);

using var watermark = new ImageWatermark(imagePath)
{
    HorizontalAlignment = HorizontalAlignment.Right,
    VerticalAlignment = VerticalAlignment.Bottom,
    SizingType = SizingType.ScaleToParentDimensions,
    ScaleFactor = 0.25,
    Opacity = 0.6,
};

var options = new PresentationWatermarkSlideOptions { SlideIndex = slideIndex };

watermarker.Add(watermark, options);
watermarker.Save(outputPath);
```

#### Technical Details

`ImageWatermark` holds an open file handle, so it needs its own `using` - note
the nested disposal pattern, with the image released before the presentation
is saved. `ScaleToParentDimensions` with `ScaleFactor = 0.25` sizes the logo to
a quarter of the slide, which survives a change of aspect ratio in a way that
absolute width and height do not.

**Key Components:**

- `ImageWatermark`: mark sourced from an image file; disposable
- `SizingType.ScaleToParentDimensions`: size relative to the slide
- `ScaleFactor`: the fraction of the parent dimension to occupy

**Parameters:**

- `imagePath`: path to the logo (`.png`, `.jpg`)
- `slideIndex`: target slide, zero-based

Output: the logo in the bottom-right corner of the chosen slide.

---

### Implementation: Diagonal, scaled, background watermark

The classic diagonal stamp, built from the formatting properties rather than
by positioning a shape manually.

```csharp
var loadOptions = new PresentationLoadOptions();
using var watermarker = new Watermarker(sourcePath, loadOptions);

var watermark = new TextWatermark(text, new Font("Arial", 48, FontStyle.Bold))
{
    ForegroundColor = Color.Gray,
    Opacity = 0.3,
    RotateAngle = rotateAngle,
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Center,
    SizingType = SizingType.ScaleToParentDimensions,
    ScaleFactor = 0.9,
    IsBackground = true,
};

watermarker.Add(watermark);
watermarker.Save(outputPath);
```

#### Technical Details

`IsBackground = true` sends the shape behind the slide content, so the mark
does not obscure text - the single most useful property here, and the one most
often missed. `RotateAngle` takes degrees, negative for counter-clockwise.
Combining `ScaleToParentDimensions` with `ScaleFactor = 0.9` makes the text
span 90% of the slide regardless of the font size, so the point size becomes a
starting hint rather than the final answer.

**Key Components:**

- `IsBackground`: places the watermark behind existing content
- `RotateAngle`: rotation in degrees; `-45` is the conventional diagonal
- `ForegroundColor`: with `Opacity`, controls how present the mark feels

**Parameters:**

- `rotateAngle`: degrees of rotation, e.g. `-45`

Output: a large grey diagonal mark behind the content of every slide.

---

### Implementation: Watermark that resists editing

Two properties turn a decorative mark into one that is awkward to remove.

```csharp
var loadOptions = new PresentationLoadOptions();
using var watermarker = new Watermarker(sourcePath, loadOptions);

var watermark = new TextWatermark(text, new Font("Arial", 40))
{
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Center,
    Opacity = 0.4,
    IsBackground = true,
};

var options = new PresentationWatermarkSlideOptions
{
    Name = "GroupDocs watermark",
    AlternativeText = text,
    IsLocked = true,
    ProtectWithUnreadableCharacters = true,
};

watermarker.Add(watermark, options);
watermarker.Save(outputPath);
```

#### Technical Details

`IsLocked` sets the shape's lock flag, so PowerPoint refuses to edit it in
place. `ProtectWithUnreadableCharacters` interleaves invisible characters
through the text, which defeats Find and Replace - at the cost of the text no
longer being searchable as the literal string, so do not rely on it for later
lookups. `Name` and `AlternativeText` label the shape, which is how you find
your own watermarks again programmatically.

Neither property makes a watermark unremovable. A determined user can still
delete the shape; these raise the effort from trivial to deliberate.

**Key Components:**

- `IsLocked`: forbids editing the watermark shape in PowerPoint
- `ProtectWithUnreadableCharacters`: breaks Find and Replace matching
- `Name` / `AlternativeText`: identify the shape for later retrieval

Output: a marked deck whose watermark resists casual tampering.

---

### Implementation: Watermark the images inside a slide

The strongest option in this set, and the one with the least obvious API. The
mark goes into the picture data rather than onto the slide.

```csharp
var loadOptions = new PresentationLoadOptions();
using var watermarker = new Watermarker(sourcePath, loadOptions);

var watermark = new TextWatermark(text, new Font("Arial", 12))
{
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Center,
    RotateAngle = 45,
    SizingType = SizingType.ScaleToParentDimensions,
    ScaleFactor = 1,
};

PresentationContent content = watermarker.GetContent<PresentationContent>();
WatermarkableImageCollection images = content.Slides[slideIndex].FindImages();

foreach (WatermarkableImage image in images)
{
    image.Add(watermark);
}

watermarker.Save(outputPath);
return images.Count;
```

#### Technical Details

`FindImages()` returns the pictures on one slide as a
`WatermarkableImageCollection`. Calling `Add` on a `WatermarkableImage`
rasterises the mark into the image bytes, so it travels with the picture if
somebody copies it into another document - which a slide-level shape does not.
Here `ScaleToParentDimensions` scales against the image, not the slide, so
`ScaleFactor = 1` covers the picture completely.

The method returns the count so a caller can tell the difference between
"watermarked three images" and "this slide had no pictures" - both of which
save successfully and look identical otherwise.

**Key Components:**

- `PresentationSlide.FindImages()`: the pictures on a single slide
- `WatermarkableImage.Add`: rasterises the mark into the image data
- `images.Count`: how many pictures were actually stamped

**Parameters:**

- `slideIndex`: which slide's images to watermark

Output: the number of images marked; `0` for a slide without pictures.

## Best Practices

- **Choose the scope before the styling.** Deck-wide, one slide, or inside the
  images are three different guarantees; the font and colour are details by
  comparison.
- **Prefer relative sizing.** `ScaleToParentDimensions` keeps a mark correct
  across 4:3 and 16:9 decks where absolute points drift.
- **Send marks to the background.** `IsBackground = true` avoids covering the
  content you are trying to protect.
- **Validate the slide index.** Read `PresentationContent.Slides.Count` first
  when the index is not a constant.
- **Treat locking as deterrence, not security.** For a guarantee that survives
  a determined recipient, watermark the embedded images or export to a format
  that is not editable.
- **Check the return of `FindImages`.** A zero count means the slide had no
  pictures, not that the operation failed.

## Additional Resources

- [Choosing a PowerPoint Watermark Strategy - 4 Approaches Compared](https://docs.groupdocs.com/watermark/net/powerpoint-watermark-scope-and-protection/)
- [Watermarking PowerPoint Slides in C#: Which Slides, and How Hard to Remove](https://blog.groupdocs.com/watermark/powerpoint-watermark-scope-and-protection-net/)
- [Adding text watermarks](https://docs.groupdocs.com/watermark/net/adding-text-watermarks/)
- [Adding repeated watermarks](https://docs.groupdocs.com/watermark/net/adding-watermark-to-images-inside-a-document/)
- [GroupDocs.Watermark Overview](https://docs.groupdocs.com/watermark/net/groupdocs-watermark-overview/)
- [API Reference](https://reference.groupdocs.com/watermark/net/)
- [Release Notes](https://releases.groupdocs.com/watermark/net/release-notes/latest/)
- [Watermark PowerPoint Presentations (blog)](https://blog.groupdocs.com/watermark/watermark-presentations/)

## Keywords

`powerpoint watermark`, `pptx watermark`, `add watermark to powerpoint c#`,
`groupdocs watermark`, `groupdocs.watermark for .net`, `presentation
watermark`, `slide watermark`, `text watermark`, `image watermark`,
`watermark specific slide`, `slideindex`, `locked watermark`, `tamper
resistant watermark`, `protectwithunreadablecharacters`, `findimages`,
`embedded image watermark`, `scaletoparentdimensions`, `document security`,
`presentation security`, `csharp`, `dotnet`, `net8`, `office open xml`,
`document watermarking api`

## Support

For technical support, visit:

- [Free Support Forum](https://forum.groupdocs.com/c/watermark/19)
- [Product Documentation](https://docs.groupdocs.com/watermark/net/)
- [Get Temporary License](https://purchase.groupdocs.com/temp-license/100354)
