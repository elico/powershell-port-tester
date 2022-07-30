$Port = 10000
$IP = "127.0.0.1"
$Address = [System.Net.IPAddress]::Parse($IP) 
$Socket = New-Object System.Net.Sockets.TCPClient($Address,$Port)
$Stream = $Socket.GetStream()
$Stream.Close()
$Socket.Close()