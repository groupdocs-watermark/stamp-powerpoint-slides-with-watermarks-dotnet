// Topic: Add watermarks to a PowerPoint PPTX presentation programmatically in C#.
// Uses GroupDocs.Watermark for .NET: text and image watermarks, whole-deck and per-slide
// targeting, formatting and rotation options, tamper protection, and watermarking the
// images embedded inside a slide.

using GroupDocs.Watermark;
using GroupDocs.Watermark.Common;
using GroupDocs.Watermark.Contents.Image;
using GroupDocs.Watermark.Contents.Presentation;
using GroupDocs.Watermark.Options.Presentation;
using GroupDocs.Watermark.Watermarks;

namespace Demo.AddWatermarkToPowerPoint;

internal static class Program
{
    private const string DocsFolder = "documents";
    private const string ResultFolder = "Result";

    private static readonly string SourcePptx = Path.Combine(DocsFolder, "sample.pptx");
    private static readonly string LogoImage = Path.Combine(DocsFolder, "logo.png");

    private const string WatermarkText = "CONFIDENTIAL";

    private static int Main()
    {
        Directory.CreateDirectory(DocsFolder);
        Directory.CreateDirectory(ResultFolder);
        ApplyLicense();

        if (!File.Exists(SourcePptx))
        {
            Console.Error.WriteLine($"Missing source presentation: {Path.GetFullPath(SourcePptx)}");
            return 1;
        }

        if (!File.Exists(LogoImage))
        {
            Console.Error.WriteLine($"Missing watermark image: {Path.GetFullPath(LogoImage)}");
            return 1;
        }

        var (width, height) = GetSlideDimensions(SourcePptx);
        Console.WriteLine($"Slide size: {width} x {height} points");

        string allSlides = Path.Combine(ResultFolder, "all-slides.pptx");
        AddTextWatermarkToAllSlides(SourcePptx, allSlides, WatermarkText);
        Console.WriteLine($"Text watermark on every slide: {Path.GetFullPath(allSlides)}");

        string oneSlide = Path.Combine(ResultFolder, "single-slide.pptx");
        AddTextWatermarkToSlide(SourcePptx, oneSlide, "DRAFT", slideIndex: 0);
        Console.WriteLine($"Text watermark on slide 1 only: {Path.GetFullPath(oneSlide)}");

        string imageOnSlide = Path.Combine(ResultFolder, "image-watermark.pptx");
        AddImageWatermarkToSlide(SourcePptx, imageOnSlide, LogoImage, slideIndex: 1);
        Console.WriteLine($"Image watermark on slide 2: {Path.GetFullPath(imageOnSlide)}");

        string formatted = Path.Combine(ResultFolder, "formatted.pptx");
        AddFormattedWatermark(SourcePptx, formatted, WatermarkText, rotateAngle: -45);
        Console.WriteLine($"Rotated, centred, scaled watermark: {Path.GetFullPath(formatted)}");

        string locked = Path.Combine(ResultFolder, "protected.pptx");
        AddLockedWatermark(SourcePptx, locked, WatermarkText);
        Console.WriteLine($"Locked watermark, hard to edit in PowerPoint: {Path.GetFullPath(locked)}");

        string slideImages = Path.Combine(ResultFolder, "slide-images.pptx");
        int stamped = AddWatermarkToSlideImages(SourcePptx, slideImages, "Protected image", slideIndex: 2);
        Console.WriteLine($"Embedded images watermarked on slide 3: {stamped}");
        Console.WriteLine($"Saved: {Path.GetFullPath(slideImages)}");

        return 0;
    }

    private static void ApplyLicense()
    {
        // Point this at your .lic file to remove evaluation limits.
        // Get a free temporary licence: https://purchase.groupdocs.com/temporary-license
        const string licensePath = "REPLACE_WITH_YOUR_LICENSE_PATH";
        if (File.Exists(licensePath))
        {
            new License().SetLicense(licensePath);
            Console.WriteLine("[license] applied");
        }
        else
        {
            Console.WriteLine("[license] no licence set - running in evaluation mode");
        }
    }

    /// <summary>
    /// Reads the slide width and height of a presentation, in points.
    /// </summary>
    /// <remarks>
    /// Calls <c>GetContent&lt;PresentationContent&gt;</c> on the <see cref="Watermarker"/> and returns
    /// <see cref="PresentationContent.SlideWidth"/> and <see cref="PresentationContent.SlideHeight"/>.
    /// These are the numbers to use when positioning a watermark in absolute coordinates instead of
    /// letting GroupDocs.Watermark align and scale it relative to the slide.
    /// </remarks>
    private static (double Width, double Height) GetSlideDimensions(string sourcePath)
    {
        var loadOptions = new PresentationLoadOptions();
        using var watermarker = new Watermarker(sourcePath, loadOptions);

        PresentationContent content = watermarker.GetContent<PresentationContent>();
        return (content.SlideWidth, content.SlideHeight);
    }

    /// <summary>
    /// Adds a text watermark to every slide of a presentation.
    /// </summary>
    /// <remarks>
    /// This is the simplest form: a <see cref="TextWatermark"/> passed to <c>Add</c> with no slide
    /// options at all, which GroupDocs.Watermark applies across the whole deck. Each watermark
    /// becomes an ordinary shape on the slide, so it survives a round-trip through PowerPoint.
    /// The result is written to <paramref name="outputPath"/>; the source file is never modified.
    /// </remarks>
    private static void AddTextWatermarkToAllSlides(string sourcePath, string outputPath, string text)
    {
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
    }

    /// <summary>
    /// Adds a text watermark to one specific slide, leaving the rest of the deck untouched.
    /// </summary>
    /// <remarks>
    /// Targeting is done with <see cref="PresentationWatermarkSlideOptions.SlideIndex"/>, which is
    /// zero-based: index 0 is the first slide. Pass those options to <c>Add</c> alongside the
    /// watermark. Use this when only a title or pricing slide needs marking rather than the
    /// whole presentation.
    /// </remarks>
    private static void AddTextWatermarkToSlide(string sourcePath, string outputPath, string text, int slideIndex)
    {
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
    }

    /// <summary>
    /// Stamps an image, such as a company logo, onto one slide of a presentation.
    /// </summary>
    /// <remarks>
    /// An <see cref="ImageWatermark"/> holds an open handle on the image file, so it is disposed
    /// before the presentation is saved - hence the nested <c>using</c>. Slide targeting works the
    /// same way as for text, through <see cref="PresentationWatermarkSlideOptions.SlideIndex"/>.
    /// <see cref="Watermark.ScaleFactor"/> combined with <see cref="SizingType.ScaleToParentDimensions"/>
    /// sizes the logo as a fraction of the slide rather than in absolute points.
    /// </remarks>
    private static void AddImageWatermarkToSlide(string sourcePath, string outputPath, string imagePath, int slideIndex)
    {
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
    }

    /// <summary>
    /// Adds a diagonal, scaled, semi-transparent watermark sent to the back of every slide.
    /// </summary>
    /// <remarks>
    /// Shows the formatting surface of <see cref="TextWatermark"/>: <see cref="Watermark.RotateAngle"/>
    /// for the classic diagonal stamp, <see cref="SizingType.ScaleToParentDimensions"/> with
    /// <see cref="Watermark.ScaleFactor"/> so the text spans the slide whatever its size, a
    /// <see cref="Font"/> with bold styling, and <see cref="Watermark.ForegroundColor"/> plus
    /// <see cref="Watermark.Opacity"/> for a readable but unobtrusive mark.
    /// <see cref="PresentationWatermarkBaseSlideOptions"/> is not needed here; instead
    /// <c>IsBackground</c> on the watermark places the shape behind the existing slide content so
    /// it does not cover text.
    /// </remarks>
    private static void AddFormattedWatermark(string sourcePath, string outputPath, string text, double rotateAngle)
    {
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
    }

    /// <summary>
    /// Adds a watermark that resists casual removal and editing in PowerPoint.
    /// </summary>
    /// <remarks>
    /// Two protections come from <see cref="PresentationWatermarkSlideOptions"/>: <c>IsLocked</c>
    /// forbids editing the watermark shape in PowerPoint, and <c>ProtectWithUnreadableCharacters</c>
    /// interleaves invisible characters into the text so Find and Replace cannot match it.
    /// <c>Name</c> and <c>AlternativeText</c> label the generated shape, which makes the watermark
    /// identifiable later - useful if it has to be located and removed programmatically.
    /// Note that protected text is no longer searchable as the literal string.
    /// </remarks>
    private static void AddLockedWatermark(string sourcePath, string outputPath, string text)
    {
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
    }

    /// <summary>
    /// Watermarks the pictures embedded inside a slide, rather than the slide itself.
    /// </summary>
    /// <remarks>
    /// A watermark placed on a slide is removed by deleting one shape; a watermark burned into an
    /// embedded picture travels with that picture even if it is copied elsewhere.
    /// <see cref="PresentationSlide.FindImages"/> returns the images on the slide as a
    /// <see cref="WatermarkableImageCollection"/>, and calling <c>Add</c> on each
    /// <see cref="WatermarkableImage"/> rasterises the watermark into the image data.
    /// Returns the number of images stamped, which is zero for a slide holding no pictures.
    /// </remarks>
    private static int AddWatermarkToSlideImages(string sourcePath, string outputPath, string text, int slideIndex)
    {
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
    }
}
