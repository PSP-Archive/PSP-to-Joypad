-- PSPtoJOYPAD for PSP (LUA[client] and VB6[server])
-- by Stefano Russello (www.stefanorussello.it)

System.setCpuSpeed(333)
System.currentDirectory("System")

ipPc = "192.168.000.001"
portPc = "8081"
nConnection = 1
maxConnection = 0
xSelect = 120
ySelect = 0

green = Color.new(0, 255, 0)
offscreen = Image.createEmpty(480, 272)
offscreen:clear(Color.new(0, 0, 0))
y = 0
x = 0
ExitFlag = false

function graphicsPrint(text)
	for i = 1, string.len(text) do
		char = string.sub(text, i, i)
		if char == "\n" then
			y = y + 8
			x = 0
		elseif char ~= "\r" then
			offscreen:print(x, y, char, green)
			x = x + 8
		end
		if y >= 272 then
			y = 0
			offscreen:clear(Color.new(0,0,0))
		end
	end
	screen:blit(0, 0, offscreen)
	screen.waitVblankStart()
	screen.flip()
end

function graphicsPrintln(text)
	graphicsPrint(text .. "\n")
end

function saveIp()
	file = io.open("ip", "w")
	if file then
		file:write(ipPc)
		file:close()
	end
end

function loadIp()
	file = io.open("ip", "r")
	if file then
		ipPc = file:read()
		file:close()
	end
	if (string.len(ipPc) == 0) then ipPc = "192.168.000.001" end
end

    buttons = {}
    buttons.analogx = 0
    buttons.analogy = 0
    buttons.cross = 0
    buttons.circle = 0
    buttons.triangle = 0
    buttons.square = 0
    buttons.up = 0
    buttons.down = 0
    buttons.left = 0
    buttons.right = 0
    buttons.l1 = 0
    buttons.r1 = 0
    buttons.l2 = 0
    buttons.r2 = 0
    buttons.select = 0
    buttons.start = 0
    
    buttons.xanalog = 51
    buttons.yanalog = 174
    buttons.xcross = 403
    buttons.ycross = 140
    buttons.xcircle = 429
    buttons.ycircle = 113
    buttons.xtriangle = 404
    buttons.ytriangle = 86
    buttons.xsquare = 377
    buttons.ysquare = 113
    buttons.xup = 52
    buttons.yup = 93
    buttons.xdown = 53
    buttons.ydown = 134
    buttons.xleft = 29
    buttons.yleft = 114
    buttons.xright = 72
    buttons.yright = 115
    buttons.xl1 = 40
    buttons.yl1 = 37
    buttons.xr1 = 385
    buttons.yr1 = 38
    buttons.xl2 = 40
    buttons.yl2 = 37
    buttons.xr2 = 385
    buttons.yr2 = 38
    buttons.xselect = 325
    buttons.yselect = 217
    buttons.xstart = 352
    buttons.ystart = 217

	aXp = false
	aYp = false
	aX = 0
	aY = 0

backshell = Image.load("backshell.png")
background = Image.load("background.png")
keyAnalog = Image.load("keys/analog.png")
keyUp = Image.load("keys/up.png")
keyDown = Image.load("keys/down.png")
keyLeft = Image.load("keys/left.png")
keyRight = Image.load("keys/right.png")
keyL1 = Image.load("keys/l1.png")
keyL2 = Image.load("keys/l2.png")
keyR1 = Image.load("keys/r1.png")
keyR2 = Image.load("keys/r2.png")
keyTriangle = Image.load("keys/triangle.png")
keyCross = Image.load("keys/cross.png")
keySquare = Image.load("keys/square.png")
keyCircle = Image.load("keys/circle.png")
keyStart = Image.load("keys/start.png")
keySelect = Image.load("keys/select.png")


offscreen:blit(0, 0, backshell)

-- init WLAN and choose connection config
Wlan.init()
configs = Wlan.getConnectionConfigs()
graphicsPrintln("PSP to JOYPAD   by www.StefanoRussello.it\n-----------------------------------------\n")
graphicsPrintln("Available connections:")
graphicsPrintln("")
ySelect = y
for key, value in ipairs(configs) do
	graphicsPrintln("> " .. key .. ": " .. value)
	maxConnection = key
end

while true do
	if Controls.read():up() then
		if nConnection <= 1 then nConnection = maxConnection else nConnection = nConnection - 1 end
		System.sleep(100)
	end
	if Controls.read():down() then
		if nConnection >= maxConnection then nConnection = 1 else nConnection = nConnection + 1 end
		System.sleep(100)
	end
	
	for i = 0, 6 do
		offscreen:drawLine(i,ySelect,i,ySelect+24,Color.new(29, 29, 29))
	end
	green = Color.new(0, 255, 0)
	y = ySelect + ((nConnection-1) * 8)
	graphicsPrintln(">")
	if Controls.read():start() or Controls.read():cross() or Controls.read():circle() then break end
end

y = ySelect
for i = 0, maxConnection do
	graphicsPrintln("")
end

graphicsPrintln("Using connection n. " .. nConnection)
graphicsPrintln("\nPlease select your IP address:\n")

ySelect = y
xPosIp = 0

loadIp()

while true do
	y = ySelect
	for i = 0, 10 do
		offscreen:drawLine(24,ySelect+i,24+(string.len(ipPc)*8),ySelect+i,Color.new(29, 29, 29))
	end
	graphicsPrintln("   " .. ipPc)
	graphicsPrint("   ")
	for i = 0, xPosIp-1 do
		graphicsPrint(" ")
	end
	graphicsPrintln("^")
	
	if Controls.read():left() then
		xPosIp = xPosIp - 1
		if xPosIp < 0 then xPosIp = 14 end
		if xPosIp == 3 or xPosIp == 7 or xPosIp == 11 then xPosIp = xPosIp - 1 end
		System.sleep(50)
	end
	if Controls.read():right() then
		xPosIp = xPosIp + 1
		if xPosIp > 14 then xPosIp = 0 end
		if xPosIp == 3 or xPosIp == 7 or xPosIp == 11 then xPosIp = xPosIp + 1 end
		System.sleep(50)
	end
	
	numberIp = string.sub(ipPc,xPosIp+1,xPosIp+1)
	if Controls.read():up() then
		numberIp = numberIp + 1
		if numberIp > 9 then numberIp = 0 end
		ipPc = string.sub(ipPc,1,xPosIp) .. numberIp .. string.sub(ipPc,xPosIp+2)
		System.sleep(50)
	end
	if Controls.read():down() then
		numberIp = numberIp - 1
		if numberIp < 0 then numberIp = 9 end
		ipPc = string.sub(ipPc,1,xPosIp) .. numberIp .. string.sub(ipPc,xPosIp+2)
		System.sleep(50)
	end
	
	if Controls.read():start() or Controls.read():cross() or Controls.read():circle() then break end
end

saveIp()

Wlan.useConnectionConfig(nConnection)

graphicsPrintln("Waiting for WLAN init and determining IP address...")
while true do
	ipAddress = Wlan.getIPAddress()
	if ipAddress then break end
	System.sleep(100)
	if Controls.read():start() then break end
end
graphicsPrintln("The PSP IP address is: " .. ipAddress)
graphicsPrintln("")

graphicsPrintln("Connecting to " .. ipPc .. "...")
socket, error = Socket.connect(ipPc, portPc)
while not socket:isConnected() do
	System.sleep(100)
	if Controls.read():start() then
		break
	end
end
graphicsPrintln("Connected to " .. tostring(socket))

bytesSent = socket:send("IP" .. ipAddress)

requestCount = 0
header = ""
headerFinished = false

offscreen:clear(Color.new(255,255,255))
screen:clear(Color.new(255,255,255))


while true do	

	pad = Controls.read()
	
	if pad:up() then
		if buttons.up == 0 then
			bytesSent = socket:send("up1,")
			buttons.up = 1
		end
	else
		if buttons.up == 1 then
			bytesSent = socket:send("up0,")
			buttons.up = 0
		end
	end

	if pad:down() then
		if buttons.down == 0 then
			bytesSent = socket:send("dw1,")
			buttons.down = 1
		end
	else
		if buttons.down == 1 then
			bytesSent = socket:send("dw0,")
			buttons.down = 0
		end
	end

	if pad:left() then
		if buttons.left == 0 then
			bytesSent = socket:send("lf1,")
			buttons.left = 1
		end
	else
		if buttons.left == 1 then
			bytesSent = socket:send("lf0,")
			buttons.left = 0
		end
	end

	if pad:right() then
		if buttons.right == 0 then
			bytesSent = socket:send("rg1,")
			buttons.right = 1
		end
	else
		if buttons.right == 1 then
			bytesSent = socket:send("rg0,")
			buttons.right = 0
		end
	end

	if pad:triangle() then
		if buttons.triangle == 0 then
			bytesSent = socket:send("tr1,")
			buttons.triangle = 1
		end
	else
		if buttons.triangle == 1 then
			bytesSent = socket:send("tr0,")
			buttons.triangle = 0
		end
	end

	if pad:cross() then
		if buttons.cross == 0 then
			bytesSent = socket:send("cr1,")
			buttons.cross = 1
		end
	else
		if buttons.cross == 1 then
			bytesSent = socket:send("cr0,")
			buttons.cross = 0
		end
	end

	if pad:square() then
		if buttons.square == 0 then
			bytesSent = socket:send("sq1,")
			buttons.square = 1
		end
	else
		if buttons.square == 1 then
			bytesSent = socket:send("sq0,")
			buttons.square = 0
		end
	end

	if pad:circle() then
		if buttons.circle == 0 then
			bytesSent = socket:send("ci1,")
			buttons.circle = 1
		end
	else
		if buttons.circle == 1 then
			bytesSent = socket:send("ci0,")
			buttons.circle = 0
		end
	end

	if pad:l() and (pad:select() or pad:start()) then
		if buttons.l2 == 0 then
			bytesSent = socket:send("l21,")
			buttons.l2 = 1
		end
	else
		if buttons.l2 == 1 then
			bytesSent = socket:send("l20,")
			buttons.l2 = 0
		end
	end
	if pad:l() and ((pad:select() or pad:start()) == false) then
		if buttons.l1 == 0 then
			bytesSent = socket:send("l11,")
			buttons.l1 = 1
		end
	else
		if buttons.l1 == 1 then
			bytesSent = socket:send("l10,")
			buttons.l1 = 0
		end
	end
	
	if pad:r() and (pad:select() or pad:start()) then
		if buttons.r2 == 0 then
			bytesSent = socket:send("r21,")
			buttons.r2 = 1
		end
	else
		if buttons.r2 == 1 then
			bytesSent = socket:send("r20,")
			buttons.r2 = 0
		end
	end
	if pad:r() and ((pad:select() or pad:start()) == false) then
		if buttons.r1 == 0 then
			bytesSent = socket:send("r11,")
			buttons.r1 = 1
		end
	else
		if buttons.r1 == 1 then
			bytesSent = socket:send("r10,")
			buttons.r1 = 0
		end
	end

	if pad:select() and pad:l() == false and pad:r() == false then
		if buttons.select == 0 then
			bytesSent = socket:send("se1,")
			buttons.select = 1
		end
	else
		if buttons.select == 1 then
			bytesSent = socket:send("se0,")
			buttons.select = 0
		end
	end

	if pad:start() and pad:l() == false and pad:r() == false then
		if buttons.start == 0 then
			bytesSent = socket:send("st1,")
			buttons.start = 1
		end
	else
		if buttons.start == 1 then
			bytesSent = socket:send("st0,")
			buttons.start = 0
		end
	end

	
	aX = Controls.read():analogX()
	aY = Controls.read():analogY()
	
	if aX >= 30 or aX <= -30 then
		socket:send("aX|" .. aX .. ",")
		aXp = true
	else
		if (aXp == true) then
			socket:send("aX|0,")
			aXp = false
		end
	end
	
	if aY >= 30 or aY <= -30 then
		socket:send("aY|" .. aY .. ",")
		aYp = true
	else
		if (aYp == true) then
			socket:send("aY|0,")
			aYp = false
		end
	end
	
	if pad:note() then
	end
	if pad:home() then
	end
	if pad:hold() then
		System.usbDiskModeActivate()
	else
		System.usbDiskModeDeactivate()
	end
	
	if pad:start() and pad:select() then
		break
	end

	offscreen:blit(0, 0, background)

	if buttons.up == 1 then offscreen:blit(buttons.xup, buttons.yup, keyUp, true) end
	if buttons.down == 1 then offscreen:blit(buttons.xdown, buttons.ydown, keyDown, true) end
	if buttons.left == 1 then offscreen:blit(buttons.xleft, buttons.yleft, keyLeft, true) end
	if buttons.right == 1 then offscreen:blit(buttons.xright, buttons.yright, keyRight, true) end
	if buttons.triangle == 1 then offscreen:blit(buttons.xtriangle, buttons.ytriangle, keyTriangle, true) end
	if buttons.cross == 1 then offscreen:blit(buttons.xcross, buttons.ycross, keyCross, true) end
	if buttons.square == 1 then offscreen:blit(buttons.xsquare, buttons.ysquare, keySquare, true) end
	if buttons.circle == 1 then offscreen:blit(buttons.xcircle, buttons.ycircle, keyCircle, true) end
	if buttons.l2 == 1 then offscreen:blit(buttons.xl2, buttons.yl2, keyL2, true) end
	if buttons.l1 == 1 then offscreen:blit(buttons.xl1, buttons.yl1, keyL1, true) end
	if buttons.r2 == 1 then offscreen:blit(buttons.xr2, buttons.yr2, keyR2, true) end
	if buttons.r1 == 1 then offscreen:blit(buttons.xr1, buttons.yr1, keyR1, true) end
	if buttons.select == 1 then offscreen:blit(buttons.xselect, buttons.yselect, keySelect, true) end
	if buttons.start == 1 then offscreen:blit(buttons.xstart, buttons.ystart, keyStart, true) end
	
	offscreen:blit(buttons.xanalog + (aX/26), buttons.yanalog + (aY/26), keyAnalog, true)

	screen:blit(0, 0, offscreen)
	screen.waitVblankStart()
	screen.flip()

end