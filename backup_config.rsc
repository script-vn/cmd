/system identity
set name=SasinNEW
/ip dns
set servers=8.8.8.8,8.8.4.4,1.1.1.1 allow-remote-requests=yes
/interface pppoe-client
add name=pppoe-out1 interface=ether1 user=sasin1 password=sasin1  add-default-route=yes default-route-distance=1 use-peer-dns=no disabled=no
add name=pppoe-out2 interface=ether2 user=sasin2 password=sasin2  add-default-route=yes default-route-distance=2 use-peer-dns=no disabled=no
/interface bridge
add name=LAN protocol-mode=none
/interface bridge port
add bridge=LAN interface=ether3
add bridge=LAN interface=ether4
add bridge=LAN interface=ether5
add bridge=LAN interface=ether6
add bridge=LAN interface=ether7
add bridge=LAN interface=ether8
/ip address
add address=192.168.1.1/24 interface=LAN
/ip pool
add name=dhcp_pool0 ranges=192.168.1.20-192.168.1.199
/ip dhcp-server
add name=dhcp1 interface=LAN address-pool=dhcp_pool0 disabled=no
/ip dhcp-server network
add address=192.168.1.0/24 gateway=192.168.1.1  dns-server=8.8.8.8,8.8.4.4,1.1.1.1
/interface vlan
add name=GUEST vlan-id=10 interface=LAN
/ip address
add address=10.0.0.1/24 interface=GUEST
/ip pool
add name=dhcp_pool1 ranges=10.0.0.10-10.0.0.254
/ip dhcp-server
add name=dhcp2 interface=GUEST address-pool=dhcp_pool1 disabled=no
/ip dhcp-server network
add address=10.0.0.0/24 gateway=10.0.0.1  dns-server=8.8.8.8,8.8.4.4,1.1.1.1
/ip firewall filter
add chain=forward src-address=10.0.0.0/24 dst-address=192.168.1.0/24  action=drop comment="Block GUEST to LAN"
/ip firewall address-list
add list=WAN address=sasin.sn.mynetname.net
/ip firewall nat
add chain=srcnat out-interface=pppoe-out1 action=masquerade
add chain=srcnat out-interface=pppoe-out2 action=masquerade
/ip firewall nat
add chain=dstnat dst-address-list=WAN protocol=tcp dst-port=8001  action=dst-nat to-addresses=192.168.1.250 to-ports=8001
add chain=dstnat dst-address-list=WAN protocol=tcp dst-port=81  action=dst-nat to-addresses=192.168.1.250 to-ports=81
/ip firewall nat
add chain=srcnat src-address=192.168.1.0/24 dst-address=192.168.1.0/24  out-interface=LAN action=masquerade comment="Hairpin NAT"
/ip service disable [find name=api]
/ip service disable [find name=api-ssl]
/ip service disable [find name=ftp]
/ip service disable [find name=ssh]
/ip service disable [find name=telnet]
/ip service disable [find name=www]
/ip service disable [find name=www-ssl]
/ip service disable [find name=dns]
/interface list
add name=all-ppp
/interface list member
add list=all-ppp interface=pppoe-out1
add list=all-ppp interface=pppoe-out2
/ip firewall raw add \
chain=prerouting \
in-interface-list=all-ppp \
protocol=tcp dst-port=53 \
action=drop \
comment="RAW drop TCP 53 from all PPP"
/ip firewall raw add \
chain=prerouting \
in-interface-list=all-ppp \
protocol=udp dst-port=53 \
action=drop \
comment="RAW drop UDP 53 from all PPP"
