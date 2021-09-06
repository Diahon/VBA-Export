using Microsoft.VisualBasic;
using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel.Design.Serialization;
using System.Drawing;
using System.IO;

namespace VBA_Export
{
    public class FormEmitter : IDisposable
    {
        public VBCodeProvider Provider { get; } = new VBCodeProvider();
        public string Name { get; private set; }

        CodeDomSerializer serializer = new CodeDomSerializer();
        DesignerSerializationManager manager = new DesignerSerializationManager();

        List<IDisposable> disposables = new List<IDisposable>();

        // https://stackoverflow.com/questions/59529839/hosting-windows-forms-designer-serialize-designer-at-runtime-and-generate-c-sh
        public FormEmitter(string name)
        {
            this.Name = name;

            // Create empty Namespace
            var frmNameSpace = new CodeNamespace();

            // Create class definition for Form
            FormType = new CodeTypeDeclaration(name) { Attributes = MemberAttributes.Public };
            FormType.IsPartial = true;
            FormType.BaseTypes.Add(new CodeTypeReference(typeof(System.Windows.Forms.Form)));

            // Create InitializeComponent method definition
            InitializeComponentMethod = new CodeMemberMethod() { Name = "InitializeComponent", Attributes = MemberAttributes.Private };
            InitializeComponentMethod.CustomAttributes.Add(new CodeAttributeDeclaration(new CodeTypeReference(typeof(System.Diagnostics.DebuggerStepThroughAttribute))));
            FormType.Members.Add(InitializeComponentMethod);

            frmNameSpace.Types.Add(FormType);
            unit.Namespaces.Add(frmNameSpace);

            // Create session for designer serialization
            disposables.Add(manager.CreateSession());

            CallMethod(new CodeThisReferenceExpression(), "SuspendLayout");
        }

        public void EmitButton(string name, string text, Rectangle rect)
        {
            InitControl(name, "Button");
            SetProperty(name, "Text", text);
            SetProperty(name, "Location", rect.Location);
            SetProperty(name, "Size", rect.Size);
        }
        public void EmitTextBox(string name, Rectangle rect)
        {
            InitControl(name, "TextBox");
            SetProperty(name, "Location", rect.Location);
            SetProperty(name, "Size", rect.Size);
        }
        public void EmitLabel(string name, string text, Rectangle rect)
        {
            InitControl(name, "Label");
            SetProperty(name, "Text", text);
            SetProperty(name, "Location", rect.Location);
            SetProperty(name, "Size", rect.Size);
        }

        public void EmitFormProperties(Size size)
        {
            // Insert Comment to mark start of control initialization
            AppendStatement(new CodeCommentStatement(""));
            AppendStatement(new CodeCommentStatement(" " + Name));
            AppendStatement(new CodeCommentStatement(""));

            SetProperty(new CodeThisReferenceExpression(), "Name", Name);
            SetProperty(new CodeThisReferenceExpression(), "Text", Name);
            SetProperty(new CodeThisReferenceExpression(), "ClientSize", size);
            SetProperty(new CodeThisReferenceExpression(), "AutoScaleMode", System.Windows.Forms.AutoScaleMode.Dpi);

            CallMethod(new CodeThisReferenceExpression(), "ResumeLayout");
        }

        #region Internal Code Generation

        #region Core

        CodeCompileUnit unit = new CodeCompileUnit();
        CodeTypeDeclaration FormType;
        CodeMemberMethod InitializeComponentMethod;

        private void AppendStatement(CodeStatement statement)
        {
            InitializeComponentMethod.Statements.Add(statement);
        }
        private void CreateVar(string name, string typeName)
        {
            var field = new CodeMemberField(new CodeTypeReference(typeName), name) { Attributes = MemberAttributes.FamilyAndAssembly };
            // http://www.dotnetframework.org/default.aspx/Dotnetfx_Win7_3@5@1/Dotnetfx_Win7_3@5@1/3@5@1/DEVDIV/depot/DevDiv/releases/whidbey/NetFXspW7/ndp/fx/src/xsp/System/Web/Compilation/BaseTemplateCodeDomTreeGenerator@cs/2/BaseTemplateCodeDomTreeGenerator@cs
            field.UserData["WithEvents"] = true;
            FormType.Members.Add(field);
        }
        #endregion

        private void InitControl(string name, string typeName)
        {
            // Create Variable in type
            CreateVar(name, typeName);

            // Insert Comment to mark start of control initialization
            AppendStatement(new CodeCommentStatement(""));
            AppendStatement(new CodeCommentStatement(" " + name));
            AppendStatement(new CodeCommentStatement(""));

            // Create instance of variable
            AppendStatement(new CodeAssignStatement(new CodeVariableReferenceExpression(name), new CodeObjectCreateExpression(new CodeTypeReference(typeName))));

            // Set property "Name" to name
            SetProperty(name, "Name", name);

            // Add control to form controls
            CallMethod(new CodePropertyReferenceExpression(new CodeThisReferenceExpression(), "Controls"), "Add", new[] { new CodeVariableReferenceExpression(name) });
        }

        #region CallMethod
        private void CallMethod(CodeExpression targetObject, string methodName, params CodeExpression[] parameters)
        {
            AppendStatement(new CodeExpressionStatement(new CodeMethodInvokeExpression(new CodeMethodReferenceExpression(targetObject, methodName), parameters)));
        }
        #endregion

        #region SetProperty
        private void SetProperty(string varName, string propName, object value)
        {
            SetProperty(new CodeVariableReferenceExpression(varName), propName, value);
        }
        private void SetProperty(CodeExpression targetObject, string propName, object value)
        {
            if (value.GetType().IsPrimitive || value.GetType() == typeof(string))
                SetProperty(targetObject, propName, new CodePrimitiveExpression(value));
            else
                SetProperty(targetObject, propName, (CodeExpression)serializer.Serialize(manager, value));
        }
        private void SetProperty(CodeExpression targetObject, string propName, CodeExpression value)
        {
            AppendStatement(new CodeAssignStatement(new CodePropertyReferenceExpression(targetObject, propName), value));
        }
        #endregion

        #endregion

        public override string ToString()
        {
            string buffer = "";
            using (var writer = new StringWriter())
            {
                Provider.GenerateCodeFromCompileUnit(unit, writer, new CodeGeneratorOptions());
                buffer += writer.ToString();
            }
            return buffer;
        }

        public void Dispose()
        {
            foreach (IDisposable item in disposables)
                try
                {
                    item.Dispose();
                }
                catch { }
        }
    }
}
