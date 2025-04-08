using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using FlaUI.Core.Input;

class CoffeeTimeApp
{
    [DllImport("kernel32.dll")]
    static extern uint SetThreadExecutionState(uint esFlags);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    const uint ES_CONTINUOUS = 0x80000000;
    const uint ES_DISPLAY_REQUIRED = 0x00000002;

    const int SW_MINIMIZE = 6;
    const int SW_RESTORE = 9;

    static int lastTeamsMinute = -1;

    static void Main()
    {
        Console.Title = "Coffee Time ☕";
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.CursorVisible = false;

        Random rand = new();

        while (true)
        {
            // Keep display on
            SetThreadExecutionState(ES_CONTINUOUS | ES_DISPLAY_REQUIRED);

            MoveMouse(rand);
            AnimateCat("You're away? I'm watching the code for you...");

            Thread.Sleep(4000);

            int currentMinute = DateTime.Now.Minute;

            if (currentMinute != lastTeamsMinute)
            {
                ReFocusTeams();
                lastTeamsMinute = currentMinute;
            }
        }
    }

    private static void MoveMouse(Random rand)
    {
        var currentPosition = Mouse.Position;
        int offsetX = rand.Next(-100, 101);
        int offsetY = rand.Next(-100, 101);
        var newPosition = new Point(currentPosition.X + offsetX, currentPosition.Y + offsetY);
        Mouse.MoveTo(newPosition);
    }

    public static void ReFocusTeams()
    {
        var processes = Process.GetProcessesByName("ms-teams");

        if (processes.Length > 0)
        {
            var teamsProcess = processes[0];
            IntPtr hWnd = teamsProcess.MainWindowHandle;

            if (hWnd == IntPtr.Zero)
            {
                Console.WriteLine($"⚠️ Could not get Teams window handle: {DateTime.Now}");
                return;
            }

            // Force focus even if background
            uint currentThreadId = GetCurrentThreadId();
            uint appThreadId = GetWindowThreadProcessId(hWnd, out _);
            AttachThreadInput(appThreadId, currentThreadId, true);

            ShowWindow(hWnd, SW_RESTORE);            // Restore if minimized
            SetForegroundWindow(hWnd);               // Try to bring to front

            AttachThreadInput(appThreadId, currentThreadId, false); // Detach again

            Console.WriteLine($"✅ Microsoft Teams window brought to foreground: {DateTime.Now}");

            Thread.Sleep(2000);

            ShowWindow(hWnd, SW_MINIMIZE); // 👈 Minimize it after interaction
        }
        else
        {
            Console.WriteLine($"⚠️ Microsoft Teams is not running: {DateTime.Now}");
        }
    }

    static void AnimateCat(string message)
    {
        string[] catFrames = new[]
        {
            @"(\____/)",
            @"( • . •)",
            @"/ >🍵  "
        };

        int catTop = 0; // top line where the cat starts

        for (int i = 0; i < 5; i++)
        {
            // Move cursor to top-left to draw cat in the same spot
            Console.SetCursorPosition(0, catTop);
            Console.WriteLine("=== Coffee Time 🐱☕ ===       ");
            foreach (var line in catFrames)
            {
                Console.WriteLine(line.PadRight(Console.WindowWidth));
            }
            Console.WriteLine(); // Spacer line
            Console.WriteLine(message.PadRight(Console.WindowWidth));

            // Clear any leftover lines from previous frame
            Console.WriteLine("".PadRight(Console.WindowWidth));
            Console.WriteLine("".PadRight(Console.WindowWidth));

            Thread.Sleep(300);

            // Blink eyes
            catFrames[1] = catFrames[1] == @"( • . •)" ? @"( - . -)" : @"( • . •)";
        }

        // After animation, move cursor down so log messages continue below cat
        Console.SetCursorPosition(0, catTop + 8);
    }
}