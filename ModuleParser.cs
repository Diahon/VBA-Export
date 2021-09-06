using Microsoft.Vbe.Interop;
using Microsoft.VisualBasic;
using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.IO;

namespace VBA_Export
{
    public class ModuleParser
    {
        public VBCodeProvider Provider { get; } = new VBCodeProvider();
        public CodeModule Module { get; private set; }

        public string Name { get => Module.Name; }
        public string FileName { get => $"{Name}.vb"; }
        public string DesignerFileName { get => $"{Name}.Designer.vb"; }

        public bool IsForm { get; set; }

        public ModuleParser(CodeModule module)
        {
            this.Module = module;
        }

        public override string ToString()
        {
            if (Module.CountOfLines == 0)
                return null;

            string code = Environment.NewLine + Module.Lines[1, Module.CountOfLines] + Environment.NewLine;

            string output = $"Module {Name}" + code + "End Module";
            if (IsForm)
                output = $"Public Class {Name}" + code + "End Class";

            // Try to format code
            try
            {
                CodeCompileUnit unit;
                using (var reader = new StringReader(output))
                {
                    unit = Provider.Parse(reader);
                }
                using (var writer = new StringWriter())
                {
                    Provider.GenerateCodeFromCompileUnit(unit, writer, new CodeGeneratorOptions());
                    output = writer.ToString();
                }
            }
            catch { }

            return output;
        }
    }
}
