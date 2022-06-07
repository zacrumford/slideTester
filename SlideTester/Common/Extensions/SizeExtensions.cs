using System;
using System.Drawing;

namespace SlideTester.Common.Extensions;

public static class SizeExtensions
{
    public static int FrameWidthDivisor { get; } = 8;
    public static int FrameHeightDivisor { get; } = 2;
    
    /// <summary>
    /// Given a Size object, compute the aspect ratio and return the ratio in a new size object
    /// </summary>
    public static Size GetAspectRatio(this Size inputSize)
    {
        int gcd = inputSize.Width.GreatestCommonDivisor(inputSize.Height);

        return new Size(inputSize.Width / gcd, inputSize.Height / gcd);
    }

    /// <summary>
    /// Given an input image size and max pixels this method will compute the scaled
    /// resolution, preserving aspect ratio, fitting it at or under the max pixels.
    /// </summary>
    /// <param name="unscaledSize">the input size to scale</param>
    /// <param name="maxPixels">the max pixel count we scale to</param>
    /// <returns>the scaled resolution</returns>
    public static Size ComputeScaledDimensions(
        this Size unscaledSize,
        int maxPixels) => unscaledSize.ComputeScaledDimensions((uint) maxPixels);
    
    /// <summary>
    /// Given an input image size and max pixels this method will compute the scaled
    /// resolution, preserving aspect ratio, fitting it at or under the max pixels.
    /// </summary>
    /// <param name="unscaledSize">the input size to scale</param>
    /// <param name="maxPixels">the max pixel count we scale to</param>
    /// <returns>the scaled resolution</returns>
    public static Size ComputeScaledDimensions(
        this Size unscaledSize,
        uint maxPixels )
    {
        ChkArg.IsLessThan( 0, unscaledSize.Width * unscaledSize.Height, "unscaledSize" );
        ChkArg.IsLessThan( ( uint )0, maxPixels, "maxPixels" );

        double scaleFactor = Math.Sqrt( maxPixels / ( double )( unscaledSize.Width * unscaledSize.Height ) );

        //
        // Now that we have the scaling factor we need to get our width and height. Compute those.
        //
        Size scaledSize = new Size( ( int )Math.Floor( unscaledSize.Width  * scaleFactor ),
            ( int )Math.Floor( unscaledSize.Height * scaleFactor ) );

        //
        // Ok, now there is a chance where due to dumb rounding issues we did not keep
        // the same aspect ratio. Check to see if that is the case. If so then fix now.
        // Additionally, our width needs to be a multiple of 8 (FrameWidthDivisor) and our
        // height needs to be a multiple of 2 (FrameHeightDivisor) check that
        // now, if it is not then the algorithm in the below block will fix it up.
        //
        if (    unscaledSize.Width * scaledSize.Height != scaledSize.Width * unscaledSize.Height
             || scaledSize.Width  % FrameWidthDivisor  != 0
             || scaledSize.Height % FrameHeightDivisor != 0 )
        {
            Size aspectRatio = unscaledSize.GetAspectRatio();

            //
            // This is kind of odd so pay attention!
            // We need to account for the scenario where one dimension (the larger) has made it all the way to
            // the next multiple of the aspect ratio but the other dimension has not. In this case we need to
            // dimension down to the previous AR multiple. For example:
            //     Unscaled dimensions: (528,352)  - 4:3
            //     max pixels: 345600
            //     scaling from logic above: (720, 479)
            // The scaled width is a multiple of 4 (the numerator of the aspect ratio, but the height
            // is not a multiple of 3 so we need to reduce the width by the AR numerator and the height
            // by height % AR denominator. If both are not multiples of the AR then reduce them both by
            // the % of the AR.
            //
            scaledSize.Width  -= scaledSize.Width  % aspectRatio.Width;
            scaledSize.Height -= scaledSize.Height % aspectRatio.Height;
            if ( scaledSize.Width * unscaledSize.Height > unscaledSize.Width * scaledSize.Height )
            {
                scaledSize.Width -= aspectRatio.Width;
            }
            else if ( scaledSize.Width * unscaledSize.Height < unscaledSize.Width * scaledSize.Height )
            {
                scaledSize.Height -= aspectRatio.Height;
            }

            //
            // Make sure that our resolution is divisible by 8, this is important because some filters
            // will not accept resolutions with widths that are not multiples of 8 (Constants.FrameWidthDivisor)
            // and heights that are not multiples of 2 (FrameHeightDivisor).
            // Exit loop if less than zero so we don't keep looking for an acceptable resolution out of our valid bounds
            //
            while (  (    scaledSize.Width  % FrameWidthDivisor  != 0
                       || scaledSize.Height % FrameHeightDivisor != 0 )
                    && scaledSize.Width  >  0
                    && scaledSize.Height >  0 )
            {
                scaledSize.Width  -= aspectRatio.Width;
                scaledSize.Height -= aspectRatio.Height;
            }
        }

        //
        // We could not find an aspect ratio that scaled (with width that was a factor of 8
        // and height factor of 2) so just nuke the aspect ratio and find something kind of close.
        // Also, scaled size may become (0,0) if the aspect ratio is odd and GCD is just 1.
        //
        if ( 0 >= scaledSize.Width || 0 >= scaledSize.Height )
        {
            scaledSize = new Size( ( int )Math.Floor( unscaledSize.Width  * scaleFactor ),
                                   ( int )Math.Floor( unscaledSize.Height * scaleFactor ) );
            scaledSize.Width  -= scaledSize.Width  % FrameWidthDivisor;
            scaledSize.Height -= scaledSize.Height % FrameHeightDivisor;
        }

        return scaledSize;
    }
}