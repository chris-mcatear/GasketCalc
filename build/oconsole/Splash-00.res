tcl86t.dll      tk86t.dll       _tk_data               y     �   H_tk_data\license.terms _tk_data\ttk\ttk.tcl zlib1.dll _tk_data\text.tcl tcl86t.dll _tk_data\ttk\fonts.tcl VCRUNTIME140.dll _tk_data\ttk\utils.tcl tk86t.dll _tk_data\tk.tcl _tk_data\ttk\cursors.tcl proc _ipc_server {channel clientaddr clientport} {
set client_name [format <%s:%d> $clientaddr $clientport]
chan configure $channel \
-buffering none \
-encoding utf-8 \
-eofchar \x04 \
-translation cr
chan event $channel readable [list _ipc_caller $channel $client_name]
}
proc _ipc_caller {channel client_name} {
chan gets $channel cmd
if {[chan eof $channel]} {
chan close $channel
exit
} elseif {![chan blocked $channel]} {
if {[string match "update_text*" $cmd]} {
global status_text
set first [expr {[string first "(" $cmd] + 1}]
set last [expr {[string last ")" $cmd] - 1}]
set status_text [string range $cmd $first $last]
}
}
}
set server_socket [socket -server _ipc_server -myaddr localhost 0]
set server_port [fconfigure $server_socket -sockname]
set env(_PYI_SPLASH_IPC) [lindex $server_port 2]
image create photo splash_image
splash_image put $_image_data
unset _image_data
proc canvas_text_update {canvas tag _var - -} {
upvar $_var var
$canvas itemconfigure $tag -text $var
}
package require Tk
set image_width [image width splash_image]
set image_height [image height splash_image]
set display_width [winfo screenwidth .]
set display_height [winfo screenheight .]
set x_position [expr {int(0.5*($display_width - $image_width))}]
set y_position [expr {int(0.5*($display_height - $image_height))}]
frame .root
canvas .root.canvas \
-width $image_width \
-height $image_height \
-borderwidth 0 \
-highlightthickness 0
.root.canvas create image \
[expr {$image_width / 2}] \
[expr {$image_height / 2}] \
-image splash_image
wm attributes . -transparentcolor magenta
.root.canvas configure -background magenta
pack .root
grid .root.canvas -column 0 -row 0 -columnspan 1 -rowspan 2
wm overrideredirect . 1
wm geometry . +${x_position}+${y_position}
wm attributes . -topmost 1
raise .�PNG

   IHDR           szz�  @IDATx��WMhA�f�m~ڦM�Զmi/�����҃Ћ�7A�&�Z0=To�E<UTězQ�"�(�`)��M�i~���$�M��M�6��ݙ}�͛o�{�f����!BD�� �AA@(�n���W!Wmq��S���e�,�8" z� /0&��$�U�Zn��T�K��*)�l���_{� ��+����b�H�K<���+�/+V���ֲͫ�ĸյ;��(9�r�]lk�U!��u
G���Ŋ�}�n "�d𷱍 
@�`�?،�a����ݳ>  �:�ù{��wȇ��{Z�"c���L�=��x�m���t	��ˢ�  �İ�t+�_jG ���7�����b�W%������8x>�f�5ۗWK��3�BP�n0�*��5�14ځ=#-���fB�Ó4�9��Q=C@��6p�Z���VE+ү�YbK��N�T\ �t����J��c�P}���G���HfQ( *79��}�/ƐJj��GNG2<�p2�e���ϫ<�|;�O�y@���$oKN)���h7�jc������A��{��n��c��:3�M<��� ��X{2��tR �z�z����B��#�{��T$����f7� ��v8`4��\�5	��.����E�a*q�^�H �^�G�dL���H���rY�l}�"������syL�'�q�T�-'W�]C94�*�Uɲ;xB�~�[�W���6����|� J!H��Ie0}=���UzW1����P��јڊrGpO��k �&�"	@�\�5�y�����g�?@��$����o��9�x�u��s�7�r� �6��x�c� �n�Ç�a`    IEND�B`�