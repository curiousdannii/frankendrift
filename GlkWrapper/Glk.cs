using System.Runtime.InteropServices;
using System.Text;

namespace GlkWrapper
{
    // Opaque Glk types
    public unsafe struct GlkFileRef { };
    public unsafe struct GlkSchannel { };
    public unsafe struct GlkStream { };
    public unsafe struct GlkWindow { };

    // Glk structs

    [StructLayout(LayoutKind.Sequential)]
    public struct Date
    {
        public uint year, month, day, weekday, hour, minute, second, microsec;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct Event
    {
        public uint type;
        public GlkWindow? win;
        public uint val1, val2;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct StreamResult
    {
        public uint readcount, writecount;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct Time
    {
        public uint high_sec, low_sec, microsec;
    }

    public class GlkWrapper
    {
        private const string DEFAULT_DLL_NAME = "Glk";
        private string DLL_NAME = DEFAULT_DLL_NAME;

        /// <summary>
        /// Set the DLL name of the Glk implementation
        /// </summary>
        /// <param name="dll_name"></param>
        public void SetGlkDllName(string dll_name)
        {
            Console.WriteLine("SetGlkDllName current: {0}, new: {1}", DLL_NAME, dll_name);
            NativeLibrary.SetDllImportResolver(typeof(GlkWrapper).Assembly, (name, asm, search) =>
            {
                Console.WriteLine("Import name {0}", name);
                if (name == DLL_NAME)
                {
                    return NativeLibrary.Load(dll_name, asm, search);
                }
                return IntPtr.Zero;
            });
            DLL_NAME = dll_name;
        }

        // And now the Glk functions
#pragma warning disable CA1401 // P/Invokes should not be visible
#pragma warning disable IDE1006 // Naming Styles

        [DllImport(DEFAULT_DLL_NAME)]
        public static extern uint glk_buffer_canon_decompose_uni(uint[] buf, uint len, uint numchars);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern uint glk_buffer_canon_normalize_uni(uint[] buf, uint len, uint numchars);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern uint glk_buffer_to_lower_case_uni(uint[] buf, uint len, uint numchars);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern uint glk_buffer_to_upper_case_uni(uint[] buf, uint len, uint numchars);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern uint glk_buffer_to_title_case_uni(uint[] buf, uint len, uint numchars, uint lowerrest);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_cancel_char_event(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_cancel_hyperlink_event(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_cancel_line_event(GlkWindow* win, Event* ev);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_cancel_mouse_event(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern byte glk_char_to_lower(byte ch);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern byte glk_char_to_upper(byte ch);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern int glk_current_simple_time(uint factor);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_current_time(Time* time);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe int glk_date_to_simple_time_local(Date* date, uint factor);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe int glk_date_to_simple_time_utc(Date* date, uint factor);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_date_to_time_local(Date* date, Time* time);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_date_to_time_utc(Date* date, Time* time);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_exit();
        [DllImport(DEFAULT_DLL_NAME)]
        private static extern unsafe GlkFileRef?* _glk_fileref_create_by_name(uint usage, byte[] name, uint rock);
        public static unsafe GlkFileRef?* glk_fileref_create_by_name(uint usage, string name, uint rock) => _glk_fileref_create_by_name(usage, StringToCString(name), rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkFileRef?* glk_fileref_create_by_prompt(uint usage, uint fmode, uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkFileRef?* glk_fileref_create_from_fileref(uint usage, GlkFileRef* fref, uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkFileRef?* glk_fileref_create_temp(uint usage, uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_fileref_delete_file(GlkFileRef* fref);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_fileref_destroy(GlkFileRef* fref);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_fileref_does_file_exist(GlkFileRef* fref);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_fileref_get_rock(GlkFileRef* fref);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkFileRef?* glk_fileref_iterate(GlkFileRef?* fref, out uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern uint glk_gestalt(uint sel, uint val);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern uint glk_gestalt_ext(uint sel, uint val, uint[] arr, uint arrlen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe int glk_get_buffer_stream(GlkStream* str, byte[] buf, uint buflen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe int glk_get_buffer_stream_uni(GlkStream* str, uint[] buf, uint buflen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe int glk_get_char_stream(GlkStream* str);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe int glk_get_char_stream_uni(GlkStream* str);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe int glk_get_line_stream(GlkStream* str, byte[] buf, uint buflen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe int glk_get_line_stream_uni(GlkStream* str, uint[] buf, uint buflen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_image_draw(GlkWindow* win, uint image, int val1, int val2);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_image_draw_scaled(GlkWindow* win, uint image, int val1, int val2, uint width, uint height);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern uint glk_image_get_info(uint image, out uint width, out uint height);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_main();
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_put_buffer(byte[] buf, uint buflen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_put_buffer_stream(GlkStream* str, byte[] buf, uint buflen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_put_buffer_stream_uni(GlkStream* str, uint[] buf, uint buflen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_put_buffer_uni(uint[] buf, uint buflen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_put_char(byte ch);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_put_char_stream(GlkStream* str, byte ch);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_put_char_stream_uni(GlkStream* str, uint ch);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_put_char_uni(uint ch);
        [DllImport(DEFAULT_DLL_NAME)]
        private static extern void _glk_put_string(byte[] s);
        public static void glk_put_string(string s) => _glk_put_string(StringToCString(s));
        [DllImport(DEFAULT_DLL_NAME)]
        private static extern unsafe void _glk_put_string_stream(GlkStream* str, byte[] s);
        public static unsafe void glk_put_string_stream(GlkStream* str, string s) => _glk_put_string_stream(str, StringToCString(s));
        [DllImport(DEFAULT_DLL_NAME)]
        private static extern unsafe void _glk_put_string_stream_uni(GlkStream* str, uint[] s);
        public static unsafe void glk_put_string_stream_uni(GlkStream* str, string s) => _glk_put_string_stream_uni(str, StringToCStringUni(s));
        [DllImport(DEFAULT_DLL_NAME)]
        private static extern void _glk_put_string_uni(uint[] s);
        public static void glk_put_string_uni(string s) => _glk_put_string_uni(StringToCStringUni(s));
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_request_char_event(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_request_char_event_uni(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_request_hyperlink_event(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_request_line_event(GlkWindow* win, char[] buf, uint maxlen, uint initlen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_request_line_event_uni(GlkWindow* win, uint[] buf, uint maxlen, uint initlen);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_request_mouse_event(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_request_timer_events(uint millisecs);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkSchannel?* glk_schannel_create(uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkSchannel?* glk_schannel_create_ext(uint rock, uint volume);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_schannel_destroy(GlkSchannel* chan);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_schannel_get_rock(GlkSchannel* chan);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkSchannel?* glk_schannel_iterate(GlkSchannel?* str, out uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_schannel_pause(GlkSchannel* chan);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_schannel_play(GlkSchannel* chan, uint sound);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_schannel_play_ext(GlkSchannel* chan, uint sound, uint repeats, uint notify);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_schannel_play_multi(GlkSchannel[] chan, uint chancount, uint[] sounds, uint soundscount, uint notify);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_schannel_set_volume(GlkSchannel* chan, uint vol);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_schannel_set_volume_ext(GlkSchannel* chan, uint vol, uint duration, uint notify);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_schannel_stop(GlkSchannel* chan);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_schannel_unpause(GlkSchannel* chan);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_select(Event* ev);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_select_poll(Event* ev);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_set_echo_line_event(GlkWindow* win, uint val);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_set_hyperlink(uint val);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_set_hyperlink_stream(GlkStream* str, uint val);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_set_style(uint styl);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_set_style_stream(GlkStream* str, uint styl);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_set_terminators_line_event(GlkWindow* win, uint[] keycodes, uint count);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_set_window(GlkWindow?* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_time_to_date_local(int time, uint factor, Date* date);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_time_to_date_utc(int time, uint factor, Date* date);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_sound_load_hint(uint sound, uint flag);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_stream_close(GlkStream* str, StreamResult?* result);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkStream?* glk_stream_get_current();
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_stream_get_position(GlkStream* str);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_stream_get_rock(GlkStream* str);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkStream?* glk_stream_iterate(GlkStream?* str, out uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkStream?* glk_stream_open_file(GlkFileRef* fileref, uint fmode, uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkStream?* glk_stream_open_file_uni(GlkFileRef* fileref, uint fmode, uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkStream?* glk_stream_open_memory(byte[]? buf, uint buflen, uint fmode, uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkStream?* glk_stream_open_memory_uni(uint[]? buf, uint buflen, uint fmode, uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkStream?* glk_stream_open_resource(uint filenum, uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkStream?* glk_stream_open_resource_uni(uint filenum, uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_stream_set_current(GlkStream?* str);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_stream_set_position(GlkStream* str, int post, uint seekmode);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_style_distinguish(GlkWindow* win, uint style1, uint style2);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_style_measure(GlkWindow* win, uint style, uint hint, out uint result);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_stylehint_clear(uint wintype, uint style, uint hint);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_stylehint_set(uint wintype, uint style, uint hint, int val);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern void glk_tick();
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_time_to_date_local(Time* time, Date* date);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_time_to_date_utc(Time* time, Date* date);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_clear(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_close(GlkWindow* win, StreamResult?* result);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_erase_rect(GlkWindow?* win, int left, int top, uint width, uint height);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_fill_rect(GlkWindow?* win, uint color, int left, int top, uint width, uint height);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_flow_break(GlkWindow?* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_get_arrangement(GlkWindow* win, out uint? method, out uint? size, out GlkWindow?* keywin);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkStream?* glk_window_get_echo_stream(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkWindow?* glk_window_get_parent(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_window_get_rock(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkWindow?* glk_window_get_sibling(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkStream?* glk_window_get_stream(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe uint glk_window_get_type(GlkWindow* win);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkWindow?* glk_window_get_root();
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_get_size(GlkWindow* win, out uint? width, out uint? height);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkWindow?* glk_window_iterate(GlkWindow?* win, out uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_move_cursor(GlkWindow* win, uint xpos, uint ypos);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe GlkWindow?* glk_window_open(GlkWindow?* splitwin, uint method, uint size, uint wintype, uint rock);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_set_arrangement(GlkWindow* win, uint method, uint size, GlkWindow?* keywin);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_set_background_color(GlkWindow?* win, uint color);
        [DllImport(DEFAULT_DLL_NAME)]
        public static extern unsafe void glk_window_set_echo_stream(GlkWindow* win, GlkStream?* stream);

#pragma warning restore CA1401 // P/Invokes should not be visible
#pragma warning restore IDE1006 // Naming Styles

        // Helpers

        private static byte[] StringToCString(string val) => Encoding.Latin1.GetBytes(val + char.MinValue);
        private static uint[] StringToCStringUni(string val) => val.EnumerateRunes().Select(r => (uint)r.Value).ToArray();
    }
}