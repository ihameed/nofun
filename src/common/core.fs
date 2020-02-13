#light "off"
(*
Copyright (c) 2015-2020, Imran Hameed
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this
  list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice,
  this list of conditions and the following disclaimer in the documentation
  and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*)
module NofunCore

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

open Microsoft.Internal.VisualStudio.Shell.Interop

open System.Runtime.InteropServices
[<DllImport("kernel32.dll")>]
extern void OutputDebugString(string lpOutputString);

//let log fmt = Printf.ksprintf System.Diagnostics.Trace.WriteLine fmt
let log fmt = Printf.ksprintf OutputDebugString fmt

let ppr fmt = Printf.ksprintf id fmt

type eol = Lf | CrLf | Cr

type settings = {
  zoom_control : bool option;
  wheel_zoom : bool option;
  suggestion : bool option;
  outlining_margin : bool option;
  eol : eol option;
  keep_eol : bool option;
  feedback : bool;
  sign_in : bool;
  notifications : bool;
}

type nofun_options = interface end

(* only available in 2015 *)
let suggestion_margin_id = new (bool EditorOptionKey) "TextViewHost/SuggestionMargin"

let code_of_eol x = match x with Lf -> "\n" | CrLf -> "\r\n" | Cr -> "\r"

let lookup (props : Utilities.PropertyCollection) k =
  let mutable ret = Unchecked.defaultof<_> in
  if props.TryGetProperty (k, &ret) then Some ret else None

let interpose_settings (mk : IEditorOptionsFactoryService) (view : IWpfTextView) =
  let props = view.Properties in
  match lookup props (typeof<nofun_options>) with
  | Some opts -> opts
  | None ->
    let old = view.Options in
    let opts = mk.CreateOptions () in
    opts.Parent <- old.Parent;
    old.Parent <- opts;
    props.AddProperty (typeof<nofun_options>, opts);
    opts

let alter_text_view cfg mk (view : IWpfTextView) =
  let edopts = view.Options in
  let nfopts = interpose_settings mk view in
  let alter' (opts : IEditorOptions) f v (k : 'a EditorOptionKey) =
    try match v with
      | Some v -> opts.SetOptionValue (k, f v)
      | None -> ignore (opts.ClearOptionValue k)
    (* raised when attempting to set a key that hasn't been registered;
       used to lazily deal with backwards compatibility for
       TextViewHost/SuggestionMargin *)
    with :? System.ArgumentException -> ()
  in
  let alter = alter' edopts id in
  alter cfg.suggestion suggestion_margin_id;
  alter cfg.outlining_margin DefaultTextViewHostOptions.OutliningMarginId;
  alter cfg.zoom_control DefaultTextViewHostOptions.ZoomControlId;
  let alter = alter' nfopts id in
  alter cfg.wheel_zoom DefaultWpfViewOptions.EnableMouseWheelZoomId;
  alter cfg.keep_eol DefaultOptions.ReplicateNewLineCharacterOptionId;
  let alter = alter' nfopts code_of_eol in
  alter cfg.eol DefaultOptions.NewLineCharacterOptionId;

type ctl = System.Windows.FrameworkElement
type shell_controls = {
  feedback : ctl option;
  sign_in : ctl option;
  notifications : ctl option;
}

let alter_controls (controls : shell_controls) (cfg : settings) =
  let alter e (c : ctl option) = match c with
    | Some c ->
      c.Width <- if e then nan else 0.;
      c.Height <- if e then nan else 0.
    | None -> () in
  alter cfg.feedback controls.feedback;
  alter cfg.sign_in controls.sign_in;
  alter cfg.notifications controls.notifications

let mutable controls = { feedback = None; sign_in = None; notifications = None }

let init_controls (cfg : settings) =
  let matches = new ((_, _) Dictionary) () in
  let add x y = ignore (matches.Add (y, x)) in
  let lookup_remove (x : ctl) =
    let k = x.GetType().FullName in
    let mutable v = Unchecked.defaultof<_> in
    //(if k.ToLower().Contains "draggable" then (log "possibly related: %A" k) else ());
    if matches.TryGetValue (k, &v) then (ignore (matches.Remove k); Some v) else None
  in
  let is_empty () = matches.Count = 0 in
  let dte = Package.GetGlobalService (typeof<EnvDTE.DTE>) :?> EnvDTE.DTE in

  add
    (fun x -> controls <- { controls with feedback = x })
    "Microsoft.VisualStudio.Feedback.SendASmileButton";
  add
    (fun x -> controls <- { controls with sign_in = x })
    "Microsoft.VisualStudio.Shell.Connected.UserInformation.UserInformationCard";

  // log "layout updated version = %A" dte.Version;
  (if dte.Version.StartsWith "16."
  then (
    // log "adding Microsoft.VisualStudio.PlatformUI.DraggableDockPanel";
    add
      (fun x -> controls <- { controls with notifications = x })
      "Microsoft.VisualStudio.PlatformUI.DraggableDockPanel"
  ) else (
    // log "Microsoft.VisualStudio.Services.UserNotifications.UserNotificationsBadge";
    add
      (fun x -> controls <- { controls with notifications = x })
      "Microsoft.VisualStudio.Services.UserNotifications.UserNotificationsBadge"
  ) );

  let wnd = Application.Current.MainWindow in
  let walk_tree () = for x in wnd.FindDescendants<ctl> () do
    begin match lookup_remove x with
    | Some save -> save (Some x)
    | None -> ()
    end;
    alter_controls controls cfg
  done in

  let mutable attempt = 0 in
  let poll_until_done () =
    let evt = wnd.LayoutUpdated in
    let rec layout_updated = System.EventHandler (fun _ _ ->
      // log "layout updated";
      walk_tree ();
      attempt <- attempt + 1;
      let should_stop = is_empty () || attempt > 400 in
      let should_stop = is_empty () in
      if should_stop then evt.RemoveHandler layout_updated) in
    evt.AddHandler layout_updated in

  walk_tree ();
  if not (is_empty ()) then poll_until_done ()

let failed x = x <> VSConstants.S_OK

let document_frames (shell : IVsUIShell) =
  let mutable frames = Unchecked.defaultof<_> in
  let res = shell.GetDocumentWindowEnum (&frames) in
  if failed res then Seq.empty else
  seq {
    let mutable cont = true in
    while cont do
      let mutable count = 0u in
      let arr = Array.zeroCreate 40 in
      let res = frames.Next (uint32 arr.Length, arr, &count) in
      if failed res || count = 0u then cont <- false
      else for x in Seq.take (int count) arr do yield x done
    done
  }

let null_to_none x = match x with
| null -> None
| x -> Some x

let mutable settings_changed : settings -> unit = fun _ -> ()

type Category = System.ComponentModel.CategoryAttribute
type Description = System.ComponentModel.DescriptionAttribute
type DisplayName = System.ComponentModel.DisplayNameAttribute
type ContentType = Microsoft.VisualStudio.Utilities.ContentTypeAttribute
type Guid = System.Runtime.InteropServices.GuidAttribute
type UiContext = Microsoft.VisualStudio.VSConstants.UICONTEXT

type StandardValuesCollection = System.ComponentModel.TypeConverter.StandardValuesCollection
type TypeConverter = System.ComponentModel.TypeConverter
type TypeConverterAttribute = System.ComponentModel.TypeConverterAttribute

type str_conv<'a when 'a : equality> (assoc) = class
  inherit TypeConverter ()
  let stdvals = Seq.toArray (Seq.map (fun (_, v) -> v) assoc)
  override x.CanConvertFrom (_, tag) = tag = typeof<string>
  override x.CanConvertTo (_, tag) = tag = typeof<string>
  override x.ConvertFrom (_, _, str) =
    let str = str :?> string in
    match Seq.tryFind (fun (k, _) -> k = str) assoc with
    | Some (_, ret : 'a) -> ret :> obj
    | None -> failwith (ppr "str_conv: couldn't parse '%A'; expected one of %A" str assoc)
  override x.ConvertTo (_, _, data, tag) =
    if tag <> typeof<string> then failwith "str_conv: expected string";
    match data with
    | :? 'a as data -> (
      match Seq.tryFind (fun (_, v) -> v = data) assoc with
      | Some (ret : string, _) -> ret :> obj
      | None -> failwith (ppr "str_conv: couldn't print '%A'; expected one of %A" data stdvals))
    | _ -> failwith "str_conv: got incorrectly-tagged data"
  override x.GetStandardValuesExclusive _ = true
  override x.GetStandardValuesSupported _ = true
  override x.GetStandardValues _ = StandardValuesCollection stdvals
end

let eol_assoc = [|
  "Unix (LF)", Some Lf;
  "Windows (CRLF)", Some CrLf;
  "Mac (pre-OS X) (CR)", Some Cr;
  "Default", None;
|]

let bool_assoc = [|
  "True", Some true;
  "False", Some false;
  "Default", None;
|]

type eol_conv () = class inherit (eol option str_conv) (eol_assoc) end

type bool_conv () = class inherit (bool option str_conv) (bool_assoc) end

let str_of assoc v = let (s, _) = Seq.find (fun (_, v') -> v = v') assoc in s

let val_of assoc s = Seq.tryPick (fun (s', v) -> if s = s' then Some v else None) assoc

let str_of_bool b = if b then "True" else "False"

let bool_of_str s = match s with "True" -> Some true | "False" -> Some false | _ -> None

let with_default def opt = match opt with Some v -> v | None -> def

let mutable cfg = {
  zoom_control = None; wheel_zoom = None; suggestion = None; outlining_margin = None;
  eol = None; keep_eol = None;
  feedback = false; sign_in = false; notifications = false }

type OptionM () = class
  member inline x.Bind (v, f) = match v with Some v -> f v | None -> None
  member inline x.Return v = Some v
  member inline x.ReturnFrom v = v
end

let opt = new OptionM ()

(* silly encoding of a polymorphic function argument that otherwise requires
   full-blown higher-ranked types
   also record fields with explicitly polymorphic types don't work in f# :[
   type parse = { f : 'a. string -> (string -> 'a option) -> 'a -> 'a }
 *)
type parse = interface
  abstract f : string -> (string -> 'a option) -> 'a -> 'a
end

let parse_settings (f : parse) = let f = f.f in
  {
    zoom_control = f "enable_zoom_control" (val_of bool_assoc) None;
    wheel_zoom = f "enable_mouse_wheel_zoom" (val_of bool_assoc) None;
    suggestion = f "show_suggestion_margin" (val_of bool_assoc) None;
    outlining_margin = f "show_outlining_margin" (val_of bool_assoc) None;
    eol = f "line_endings" (val_of eol_assoc) None;
    keep_eol = f "preserve_line_endings" (val_of bool_assoc) None;
    feedback = f "show_feedback_button" bool_of_str true;
    sign_in = f "show_sign_in_button" bool_of_str true;
    notifications = f "show_notifications_button" bool_of_str true;
  }

let print_settings f (settings : settings) =
    f "enable_zoom_control" (str_of bool_assoc settings.zoom_control);
    f "enable_mouse_wheel_zoom" (str_of bool_assoc settings.wheel_zoom);
    f "show_suggestion_margin" (str_of bool_assoc settings.suggestion);
    f "show_outlining_margin" (str_of bool_assoc settings.outlining_margin);
    f "line_endings" (str_of eol_assoc settings.eol);
    f "preserve_line_endings" (str_of bool_assoc settings.keep_eol);
    f "show_feedback_button" (str_of_bool settings.feedback);
    f "show_sign_in_button" (str_of_bool settings.sign_in);
    f "show_notifications_button" (str_of_bool settings.notifications)

let registry_path = "DialogPage\Nofun+settings_page"

let load_settings_from_registry (root : Microsoft.Win32.RegistryKey) =
  let sub = root.OpenSubKey registry_path in
  if sub <> null then begin
    use sub = sub in
    let get name parse def = with_default def <| opt {
      let! v = null_to_none (sub.GetValue name) in
      let! v = match v with :? string as v -> Some v | _ -> None in
      let! v = parse v in
      return v
      } in
    parse_settings { new parse with member __.f x y z = get x y z }
  end
  else begin
    parse_settings { new parse with member __.f _ _ def = def }
  end

let save_settings_to_registry (root : Microsoft.Win32.RegistryKey) (settings : settings) =
  let sub = root.OpenSubKey (registry_path, true) in
  use sub = if sub <> null then sub else root.CreateSubKey registry_path in
  let set name v = sub.SetValue (name, v) in
  print_settings set settings

[<Guid "2880a7df-1a69-4e26-8743-b9f2a9572cd2">]
[<Sealed>]
type settings_page<'a> () = class
  inherit DialogPage ()

  [<DisplayName "Enable Mouse Wheel Zoom">]
  [<Description "Enables or disables mouse wheel zooming in text views.">]
  [<Category "Editor">]
  [<TypeConverter (typeof<bool_conv>)>]
  member val enable_mouse_wheel_zoom = None with get, set

  [<DisplayName "Enable Zoom Control">]
  [<Description "Enables or disables the zoom control in text views.">]
  [<Category "Editor">]
  [<TypeConverter (typeof<bool_conv>)>]
  member val enable_zoom_control = None with get, set

  [<DisplayName "Show Suggestion Margin">]
  [<Description "Shows or hides the suggestion margin.">]
  [<Category "Editor">]
  [<TypeConverter (typeof<bool_conv>)>]
  member val show_suggestion_margin = None with get, set

  [<DisplayName "Show Outlining Margin">]
  [<Description "Shows or hides the outlining margin.">]
  [<Category "Editor">]
  [<TypeConverter (typeof<bool_conv>)>]
  member val show_outlining_margin = None with get, set

  [<DisplayName "Line Endings">]
  [<Description "Sets the end-of-line code.">]
  [<Category "Editor">]
  [<TypeConverter (typeof<eol_conv>)>]
  member val line_endings = None : eol option with get, set

  [<DisplayName "Preserve Line Endings">]
  [<Description "If true, the editor will use the existing end of line code present in the document. Else, the editor will use the end of line code specified by the \"Line Endings\" option.">]
  [<Category "Editor">]
  [<TypeConverter (typeof<bool_conv>)>]
  member val preserve_line_endings = None : bool option with get, set

  [<DisplayName "Show Feedback Button">]
  [<Description "Shows or hides the title bar feedback button.">]
  [<Category "Shell">]
  member val show_feedback_button = true with get, set

  [<DisplayName "Show Sign In Button">]
  [<Description "Shows or hides the menu bar sign in button.">]
  [<Category "Shell">]
  member val show_sign_in_button = true with get, set

  [<DisplayName "Show Notifications Button">]
  [<Description "Shows or hides the title bar notifications button.">]
  [<Category "Shell">]
  member val show_notifications_button = true with get, set

  member x.settings () =
    let package = x.GetService (typeof<'a>) :?> Package in
    let path = x.SettingsRegistryPath in
    use root = package.UserRegistryRoot in
    // log "root = %A, path = %A" root path;
  {
    zoom_control = x.enable_zoom_control;
    wheel_zoom = x.enable_mouse_wheel_zoom;
    suggestion = x.show_suggestion_margin;
    outlining_margin = x.show_outlining_margin;
    feedback = x.show_feedback_button;
    eol = x.line_endings;
    keep_eol = x.preserve_line_endings;
    sign_in = x.show_sign_in_button;
    notifications = x.show_notifications_button;
  }

  member x.from_settings (cfg : settings) =
    x.enable_zoom_control <- cfg.zoom_control;
    x.enable_mouse_wheel_zoom <- cfg.wheel_zoom;
    x.show_suggestion_margin <- cfg.suggestion;
    x.show_outlining_margin <- cfg.outlining_margin;
    x.line_endings <- cfg.eol;
    x.preserve_line_endings <- cfg.keep_eol;
    x.show_feedback_button <- cfg.feedback;
    x.show_sign_in_button <- cfg.sign_in;
    x.show_notifications_button <- cfg.notifications;

  member x.changed () =
    let cfg = x.settings () in
    settings_changed cfg

  member x.print f =
    f "enable_zoom_control" (str_of bool_assoc x.enable_zoom_control);
    f "enable_mouse_wheel_zoom" (str_of bool_assoc x.enable_mouse_wheel_zoom);
    f "show_suggestion_margin" (str_of bool_assoc x.show_suggestion_margin);
    f "show_outlining_margin" (str_of bool_assoc x.show_outlining_margin);
    f "line_endings" (str_of eol_assoc x.line_endings);
    f "preserve_line_endings" (str_of bool_assoc x.preserve_line_endings);
    f "show_feedback_button" (str_of_bool x.show_feedback_button);
    f "show_sign_in_button" (str_of_bool x.show_sign_in_button);
    f "show_notifications_button" (str_of_bool x.show_notifications_button);

  override x.OnApply y = base.OnApply y; x.changed ()

  override x.SaveSettingsToStorage () =
    let package = x.GetService (typeof<'a>) :?> Package in
    use root = package.UserRegistryRoot in
    let settings = x.settings () in
    save_settings_to_registry root settings

  override x.LoadSettingsFromStorage () =
    let package = x.GetService (typeof<'a>) :?> Package in
    use root = package.UserRegistryRoot in
    let cfg = load_settings_from_registry root in
    x.from_settings cfg;
    x.changed ()

  override x.SaveSettingsToXml xml =
    let set name v = ignore (xml.WriteSettingString (name, v)) in
    x.print set

  override x.LoadSettingsFromXml xml =
    let read_str name =
      let mutable str = Unchecked.defaultof<_> in
      if not (failed (xml.ReadSettingString (name, &str)))
      then Some str else None in
    let get name parse def = with_default def <| opt {
      let! v = read_str name in
      let! v = parse v in
      return v
      } in
    let cfg = parse_settings { new parse with member __.f x y z = get x y z } in
    x.from_settings cfg;
    x.changed ()
end

[<Export (typeof<IWpfTextViewCreationListener>)>]
[<ContentType "text">]
[<TextViewRole (PredefinedTextViewRoles.Document)>]
[<Sealed>]
type text_view_created [<ImportingConstructor>]
  (mk : IEditorOptionsFactoryService) =
class
  interface IWpfTextViewCreationListener with
    member __.TextViewCreated view = alter_text_view cfg mk view
  end
end
