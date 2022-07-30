$port = 10000
$endpoint = new-object System.Net.IPEndPoint([ipaddress]::any,$port) 
$listener = new-object System.Net.Sockets.TcpListener $EndPoint
$listener.start()
$data = $listener.AcceptTcpClient() 
$stream = $data.GetStream() 
$stream.close()
$listener.stop()