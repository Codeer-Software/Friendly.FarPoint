using Codeer.Friendly.Dynamic;
using Codeer.Friendly.Windows;
using Codeer.Friendly.Windows.Grasp;
using Codeer.TestAssistant.GeneratorToolKit;
using Friendly.FarPoint;

namespace Test.CaptureTest
{
    [WindowDriver(TypeFullName = "Target.MainForm")]
    public class MainForm_Driver
    {
        public WindowControl Core { get; }
        public FpSpreadDriver _spread => new FpSpreadDriver(Core.Dynamic()._spread);

        public MainForm_Driver(WindowControl core)
        {
            Core = core;
        }
    }

    public static class MainForm_Driver_Extensions
    {
        [WindowDriverIdentify(TypeFullName = "Target.MainForm")]
        public static MainForm_Driver Attach_MainForm(this WindowsAppFriend app)
            => new MainForm_Driver(app.WaitForIdentifyFromTypeFullName("Target.MainForm"));
    }
}