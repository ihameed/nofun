#light "off"
module Nofun

open Microsoft.VisualStudio
open Microsoft.VisualStudio.ComponentModelHost
open Microsoft.VisualStudio.Editor
open Microsoft.VisualStudio.PlatformUI
open Microsoft.VisualStudio.Shell
open Microsoft.VisualStudio.Shell.Interop
open Microsoft.VisualStudio.Text.Editor
open System.Collections.Generic
open System.ComponentModel.Composition
open System.Windows

open NofunCore

[<PackageRegistration (UseManagedResourcesOnly = true)>]
[<InstalledProductRegistration ("#110", "#112", "1.0.0.0", IconResourceID = 400)>]
[<Guid ("e1f5e968-3b43-45e3-9c49-24220d9ff002")>]
[<ProvideOptionPage (typeof<nofun_package settings_page>, "Nofun", "General", 1s, 2s, false)>]
[<ProvideProfile (typeof<nofun_package settings_page>, "Nofun", "General", 1s, 3s, false)>]
[<ProvideAutoLoad (UiContext.NoSolution_string)>]
[<ProvideAutoLoad (UiContext.SolutionOpening_string)>]
[<ProvideAutoLoad (UiContext.SolutionExists_string)>]
[<Sealed>]
type nofun_package () = class
  inherit Package ()
  override x.Initialize () =
    base.Initialize ();

    let mef = x.GetService typeof<SComponentModel> :?> IComponentModel in
    let shell = x.GetService typeof<SVsUIShell> :?> IVsUIShell in
    let cvt = mef.GetService<IVsEditorAdaptersFactoryService> () in
    let mk = mef.GetService<IEditorOptionsFactoryService> () in
    let reconfigure () =
      document_frames shell |>
        Seq.choose (null_to_none << VsShellUtilities.GetTextView) |>
        Seq.choose (null_to_none << cvt.GetWpfTextView) |>
        Seq.filter (fun x -> x.Roles.Contains PredefinedTextViewRoles.Document) |>
        Seq.filter (fun x -> x.TextDataModel.ContentType.IsOfType "text") |>
        Seq.iter (alter_text_view cfg mk);
      alter_controls controls cfg in

    settings_changed <- (fun cfg' -> if cfg <> cfg' then
      cfg <- cfg';
      reconfigure ());
    let page = x.GetDialogPage (typeof<nofun_package settings_page>) :?> nofun_package settings_page in
    cfg <- page.settings ();
    init_controls cfg;
    reconfigure ()
end

