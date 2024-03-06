# VBA-Export
Export VBA Code and UserForms to files (partial) compatible with Windows Forms vb.net Project.   
Currently this addon is only usable in Excel but the code base is easily adaptable to other office applications (possibly by using a shared underlaying library) ...

## ToDo
 - [ ] Add `Sub New()` (Constructor) to call `InitializeComponents()`
 - [ ] Subscribe event callbacks to the .Net events
 - [ ] Compatibility Layer
    - [ ] `TextBox.Value` => `TextBox.Text`
    - [ ] Custom conversions
    - [ ] ...
