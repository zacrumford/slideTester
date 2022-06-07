using System;
using System.Collections.Generic;

namespace SlideTester.Driver
{
    public class Settings
    {
        public static Settings Default { get; set; } = new Settings();
        
        // Copied from Panopto.PackagerWorker.Settings.Default.PPTImageMaxPixelCount
        public uint SlideMaxPixelCountFallback { get; } = 1228800;

        // Note: will populate when we implement slideProcessorCore
        public List<string> CustomFontDirectories = new List<string>();

        // Copied from Panopto.PackagerWorker.Settings.Default.PPTCopyFileTimeout 
        public TimeSpan PowerpointCopyFileTimeout { get; } = TimeSpan.FromMinutes(2);
        
        // Copied from Panopto.PackagerWorker.Settings.Default.PPTForceCloseTimeout 
        public TimeSpan PowerpointForceCloseTimeout { get; } = TimeSpan.FromSeconds(30);
        
        public TimeSpan PowerpointAcquireLockTimeout { get; } = TimeSpan.FromSeconds(3);
        
        // Copied from Panopto.PackagerWorker.Settings.Default.COMRetryCount 
        public int ComRetryCount { get; } = 5;

        // Copied from Panopto.Backend.ServiceTaskWorker.PowerpointDocumentProcessor.PPTCOMTimeout
        public TimeSpan ComTimeout { get; } = TimeSpan.FromMinutes(5);
        
        // Copied from Panopto.PackagerWorker.Settings.Default.PowerpointSleepDuration
        public TimeSpan ComRetryDelay { get; } = TimeSpan.FromSeconds(1);
        
        // Copied from Panopto.PackagerWorker.Settings.Default.PPTProcessingExecutionTimeout
        public TimeSpan ProcessingTimeout { get; } = TimeSpan.FromMinutes(120);

        public int MaxSlideProcessingParallelization { get; } = 10;

        
        // BUGBUG: tfs-119165: License Aspose for Production use
        // - We should rename this var so it is not easily discoverable in the decompilation or better
        // - yet, Aspose recommends doing a simple static encryption (e.g. fixed key) and using a
        // - resource file to hold the encrypted license. 
        public string AsposeLicense { get; } =
            @"<?xml version=""1.0""?>
            <License>
                <Data>
                    <LicensedTo>Panopto</LicensedTo>
                    <EmailTo>zrumford@panopto.com</EmailTo>
                    <LicenseType>Developer OEM</LicenseType>
                    <LicenseNote>1 Developer And Unlimited Deployment Locations</LicenseNote>
                    <OrderID>220509205321</OrderID>
                    <UserID>603654</UserID>
                    <OEM>This is a redistributable license</OEM>
                    <Products>
                    <Product>Aspose.Slides for .NET</Product>
                    </Products>
                    <EditionType>Professional</EditionType>
                    <SerialNumber>c48189bf-fabe-4ef6-a59b-a67804d47a4b</SerialNumber>
                    <SubscriptionExpiry>20230525</SubscriptionExpiry>
                    <LicenseExpiry>20220625</LicenseExpiry>
                    <ExpiryNote>This is a temporary license for non-commercial use only and it will expire on 2022-06-25</ExpiryNote>
                    <LicenseVersion>3.0</LicenseVersion>
                    <LicenseInstructions>https://purchase.aspose.com/policies/use-license</LicenseInstructions>
                </Data>
                <Signature>mMVlitrPAY9r1jPEjyUSfhrpwDEv40lITOR8vTcw5A+MPkWZAPEDOtnKsHWeZG45qYOLtUxVmpmObNqiG34urmbfFE3ntWnlcIpGTgHk4pd1+SQO7e1zCY8sDnIFG19z5K0SN++O4hEfG8cS9RU+vviTMFtUFgtg040RKNlMuHEwQEEN3CHDH1wcME61wb/Ncc17BC16214Xt/563R2uphJXvp9b7pWhM2Bi7wvMbPwVV0P+ed2kGIfP9JLMe99JBURjP71GL9HKW5Q4k5YsvRACjj7pn9fXhTWZDUvI6WVEg2iyQn82MAlSOQa5mvhDz1BpCs+BmAlJFuX1mp+lNA==</Signature>
            </License>";

    }
}
