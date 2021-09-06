using Microsoft.Office.Tools.Ribbon;
using Microsoft.Vbe.Interop;
using Microsoft.Vbe.Interop.Forms;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using Label = Microsoft.Vbe.Interop.Forms.Label;
using TextBox = Microsoft.Vbe.Interop.Forms.TextBox;

namespace VBA_Export
{
    public partial class MainRibbonBar
    {
        string ProjectDirectory = null;

        private void ExportButton_Click(object sender, RibbonControlEventArgs e)
        {
            VBProject project = ThisAddIn.App.ActiveWorkbook.VBProject;

            ProjectDirectory = Path.GetFullPath(project.Name);
            if (!Directory.Exists(ProjectDirectory))
                Directory.CreateDirectory(ProjectDirectory);

            foreach (VBComponent comp in project.VBComponents)
            {
                ModuleParser module = new ModuleParser(comp.CodeModule);

                UserForm frm = comp.Designer as UserForm;
                if (frm != null)
                {
                    // Inform module parser that a partial class should be generated
                    module.IsForm = true;

                    // Generate designer code for windows forms
                    FormEmitter emitter = new FormEmitter(module.Name);
                    foreach (Control control in frm.Controls)
                    {
                        CommandButton button = control as CommandButton;
                        if (button != null)
                        {
                            emitter.EmitButton(control.Name, button.Caption, new Rectangle((int)control.Left, (int)control.Top, (int)control.Width, (int)control.Height));
                            continue;
                        }

                        Label label = control as Label;
                        if (label != null)
                        {
                            emitter.EmitLabel(control.Name, label.Caption, new Rectangle((int)control.Left, (int)control.Top, (int)control.Width, (int)control.Height));
                            continue;
                        }

                        TextBox textBox = control as TextBox;
                        if (textBox != null)
                        {
                            emitter.EmitTextBox(control.Name, new Rectangle((int)control.Left, (int)control.Top, (int)control.Width, (int)control.Height));
                            continue;
                        }
                    }
                    emitter.EmitFormProperties(new Size((int)frm.InsideWidth, (int)frm.InsideHeight));

                    // Write designer code to file
                    File.WriteAllText(Path.Combine(ProjectDirectory, module.DesignerFileName), emitter.ToString());

                    // Cleanup
                    emitter.Dispose();
                }

                // Write (module) code to file
                string code = module.ToString();
                if (!string.IsNullOrEmpty(code))
                    File.WriteAllText(Path.Combine(ProjectDirectory, module.FileName), code);
            }
        }

        private void OpenFolderButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrEmpty(ProjectDirectory))
                return;

            try
            {
                Process.Start("explorer", ProjectDirectory);
            }
            catch { }
        }
    }
}
