#!/usr/bin/python
# -*- coding: utf-8 -*-

# Reference: https://github.com/frostschutz/SourceLib/blob/master/SourceRcon.py
#------------------------------------------------------------------------------
# SourceRcon - Python class for executing commands on Source Dedicated Servers
# Copyright (c) 2010 Andreas Klauer <Andreas.Klauer@metamorpher.de>
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#------------------------------------------------------------------------------

"""http://developer.valvesoftware.com/wiki/Source_RCON_Protocol"""

import select
import socket
import struct

SERVERDATA_AUTH = 3
SERVERDATA_AUTH_RESPONSE = 2

SERVERDATA_EXECCOMMAND = 2
SERVERDATA_RESPONSE_VALUE = 0

MAX_COMMAND_LENGTH=510 # found by trial & error

MIN_MESSAGE_LENGTH=4+4+1+1 # command (4), id (4), string1 (1), string2 (1)
MAX_MESSAGE_LENGTH=4+4+4096+1 # command (4), id (4), string (4096), string2 (1)

# there is no indication if a packet was split, and they are split by lines
# instead of bytes, so even the size of split packets is somewhat random.
# Allowing for a line length of up to 400 characters, risk waiting for an
# extra packet that may never come if the previous packet was this large.
PROBABLY_SPLIT_IF_LARGER_THAN = MAX_MESSAGE_LENGTH - 400

class SourceRconError(Exception):
    pass

class SourceRcon(object):
    """Example usage:

       import SourceRcon
       server = SourceRcon.SourceRcon('1.2.3.4', 27015, 'secret')
       print server.rcon('cvarlist')
    """
    def __init__(self, host, port=27015, password='', timeout=1.0):
        self.host = host
        self.port = port
        self.password = password
        self.timeout = timeout
        self.tcp = False
        self.reqid = 0

    def disconnect(self):
        """Disconnect from the server."""
        if self.tcp:
            self.tcp.close()

    def connect(self):
        """Connect to the server. Should only be used internally."""
        try:
            self.tcp = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.tcp.settimeout(self.timeout)
            self.tcp.connect((self.host, self.port))
        except socket.error as msg:
            raise SourceRconError('Disconnected from RCON, please restart program to continue.')

    def send(self, cmd, message):
        """Send command + payload to the server (internal)."""
        if isinstance(message, str):
            message = message.encode("utf-8")  # str → bytes
        if len(message) > MAX_COMMAND_LENGTH:
            raise SourceRconError("RCON message too large to send")

        self.reqid += 1
        payload = struct.pack("<l", self.reqid)
        payload += struct.pack("<l", cmd)
        payload += message + b"\x00\x00"  # two NULs

        self.tcp.send(struct.pack("<l", len(payload)) + payload)

    def receive(self):
        """Read and return *one* logical RCON reply (handles split packets)."""
        full = b""
        response_type = None

        while True:
            # --- Wait for at least the size header ---
            try:
                header = self.tcp.recv(4)
            except socket.timeout:
                raise SourceRconError("Timed out while waiting for reply")
            if len(header) < 4:      # remote closed?
                raise SourceRconError("RCON connection closed by remote host")

            packet_len = struct.unpack("<l", header)[0]
            if packet_len < MIN_MESSAGE_LENGTH or packet_len > MAX_MESSAGE_LENGTH:
                raise SourceRconError(f"Illegal packet size: {packet_len}")

            # --- Read the rest of this packet ---
            chunk = b""
            while len(chunk) < packet_len:
                part = self.tcp.recv(packet_len - len(chunk))
                if not part:
                    raise SourceRconError("RCON connection closed while reading packet")
                chunk += part

            # --- Parse this packet ---
            req_id, resp = struct.unpack("<ll", chunk[:8])
            if req_id == -1:
                raise SourceRconError("Bad RCON password")
            if req_id != self.reqid:
                raise SourceRconError(f"Req‑ID mismatch {req_id} != {self.reqid}")

            response_type = resp
            full += chunk[8:]                     # append body (incl. NULs etc.)

            # Heuristic: if nothing else is waiting *and* this packet wasn’t huge,
            # assume we’ve got the last part of the reply.
            if not select.select([self.tcp], [], [], 0)[0] and packet_len < PROBABLY_SPLIT_IF_LARGER_THAN:
                break

        # Strip the two trailing NULs and decode once
        if full.endswith(b"\x00\x00"):
            full = full[:-2]
        return full.decode("utf-8", "ignore")

    def rcon(self, command):
        """Send RCON command to the server. Connect and auth if necessary,
           handle dropped connections, send command and return reply."""
        # special treatment for sending whole scripts
        if '\n' in command:
            commands = command.split('\n')
            def f(x): y = x.strip(); return len(y) and not y.startswith("//")
            commands = filter(f, commands)
            results = map(self.rcon, commands)
            return "".join(results)

        # send a single command. connect and auth if necessary.
        try:
            self.send(SERVERDATA_EXECCOMMAND, command)
            return self.receive()
        except:
            # timeout? invalid? we don't care. try one more time.
            self.disconnect()
            self.connect()
            self.send(SERVERDATA_AUTH, self.password)

            auth = self.receive()
            # the first packet may be a "you have been banned" or empty string.
            # in the latter case, fetch the second packet
            if auth == '':
                auth = self.receive()

            if auth is not True:
                self.disconnect()
                raise SourceRconError('RCON authentication failure: %s' % (repr(auth),))

            self.send(SERVERDATA_EXECCOMMAND, command)
            return self.receive()