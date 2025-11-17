using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime;

namespace Lab1
{
    public class ErrorListener : BaseErrorListener
    {
        public List<string> Errors { get; } = new List<string>();

        public override void SyntaxError(
            TextWriter output,
            IRecognizer recognizer,
            IToken offendingSymbol,
            int line,
            int charPositionInLine,
            string msg,
            RecognitionException e)
        {
            Errors.Add($"line {line}:{charPositionInLine} {msg}");
        }
    }
}
