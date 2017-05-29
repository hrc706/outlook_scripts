tell application "Microsoft Outlook"
	activate
	set curMsgs to current messages
	if (count of curMsgs) > 0 then
		set theMsg to item 1 of curMsgs
		if class of theMsg is (meeting message) then
			if type of theMsg is (request meeting type) then
				accept invite theMsg with sending response
			else if type of theMsg is (cancellation meeting type) then
				-- do some tricky thing
			end if
		end if
	end if
end tell
