using System;
using System.Diagnostics;

namespace MyStuff2Docx {
    class ProcessingMessage : IDisposable {
        private const string PROCESSING_MESSAGE = "Processing...";
        private const string DONE_MESSAGE = "Done";
        private const string ERROR_MESSAGE = "Error!";
        private const ConsoleColor PROCESSING_COLOR = ConsoleColor.Yellow;
        private const ConsoleColor DONE_COLOR = ConsoleColor.Green;
        private const ConsoleColor ERROR_COLOR = ConsoleColor.Red;

        private int CursorXPosition;
        private int CursorYPosition;

        public bool HasError { get; set; } = false;

        public ProcessingMessage(string text) {
            lock (Program._CONSOLE_WRITE_LOCK) {
                var currentConsoleColor = Console.ForegroundColor;

                Console.Write($"{text} --> ");
                Console.ForegroundColor = PROCESSING_COLOR;
                Console.Write($"{PROCESSING_MESSAGE}");
                Console.ForegroundColor = currentConsoleColor;

                CursorXPosition = Console.CursorLeft;
                CursorYPosition = Console.CursorTop;
                Console.WriteLine();
            }
        }

        public void Dispose() {
            lock (Program._CONSOLE_WRITE_LOCK) {
                var currentXPosition = Console.CursorLeft;
                var currentYPosition = Console.CursorTop;
                var currentConsoleColor = Console.ForegroundColor;

                try {
                    Console.SetCursorPosition(CursorXPosition - PROCESSING_MESSAGE.Length, CursorYPosition);
                }
                finally {
                    if (HasError) {
                        Console.ForegroundColor = ERROR_COLOR;
                        Console.Write(ERROR_MESSAGE.PadRight(PROCESSING_MESSAGE.Length));
                    }
                    else {
                        Console.ForegroundColor = DONE_COLOR;
                        Console.Write(DONE_MESSAGE.PadRight(PROCESSING_MESSAGE.Length));
                    }

                    Console.SetCursorPosition(currentXPosition, currentYPosition);
                    Console.ForegroundColor = currentConsoleColor;
                }
            }
        }
    }
}
