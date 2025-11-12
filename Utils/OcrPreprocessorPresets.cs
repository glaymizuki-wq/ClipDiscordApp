using OpenCvSharp;
using YourNamespace.Ocr;

namespace ClipDiscordApp.Utils
{
    public static class OcrPreprocessorPresets
    {
        public static OcrPreprocessor.Params ReadableText => new OcrPreprocessor.Params
        {
            Scale = 4,
            BilateralDiameter = 9,
            BilateralSigmaColor = 75,
            BilateralSigmaSpace = 75,
            GaussianKernel = 3,

            AdaptiveBlockSize = 31,
            AdaptiveC = 6.0,
            InvertThreshold = false,

            UseClahe = true,
            ClaheClipLimit = 3.0,
            ClaheTileGridSize = new OpenCvSharp.Size(8, 8),

            MorphKernel = 1,     // 文字の癒着を避ける
            UseDilate = false,
            DilateIterations = 0,
            Border = 8,

            Sharpen = true,
            SharpenAmount = 1.0,
            SharpenBlurCoeff = -0.1,

            IsDarkBackground = false,
            MakeOutline = false,

            SavePreprocessed = true,
            PreprocessedSaveFolder = null,
            PreprocessedFilenamePrefix = "pre_readable"
        };

        public static OcrPreprocessor.Params Default => new OcrPreprocessor.Params
        {
            Scale = 4,
            BilateralDiameter = 9,
            BilateralSigmaColor = 75,
            BilateralSigmaSpace = 75,
            GaussianKernel = 3,

            AdaptiveBlockSize = 31,
            AdaptiveC = 8.0,
            InvertThreshold = false,

            UseClahe = true,
            ClaheClipLimit = 3.0,
            ClaheTileGridSize = new OpenCvSharp.Size(8, 8),

            MorphKernel = 3,
            UseDilate = true,
            DilateIterations = 1,
            Border = 8,

            Sharpen = false,
            SharpenAmount = 1.2,
            SharpenBlurCoeff = -0.2,

            IsDarkBackground = false,
            MakeOutline = false,

            SavePreprocessed = true,
            PreprocessedSaveFolder = null,
            PreprocessedFilenamePrefix = "pre_default"
        };

        public static OcrPreprocessor.Params Conservative => new OcrPreprocessor.Params
        {
            Scale = 5,
            BilateralDiameter = 9,
            BilateralSigmaColor = 100,
            BilateralSigmaSpace = 100,
            GaussianKernel = 5,

            AdaptiveBlockSize = 41,
            AdaptiveC = 10.0,
            InvertThreshold = false,

            UseClahe = false,
            ClaheClipLimit = 3.0,
            ClaheTileGridSize = new OpenCvSharp.Size(8, 8),

            MorphKernel = 5,
            UseDilate = false,
            DilateIterations = 0,
            Border = 10,

            Sharpen = false,
            SharpenAmount = 1.0,
            SharpenBlurCoeff = -0.1,

            IsDarkBackground = false,
            MakeOutline = false,

            SavePreprocessed = true,
            PreprocessedSaveFolder = null,
            PreprocessedFilenamePrefix = "pre_conservative"
        };
    }
}