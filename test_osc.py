#!/usr/bin/env python3
"""
Test script for OSC functionality
"""

from pythonosc import udp_client
import time

def test_osc_status():
    """Test the OSC status endpoint"""
    print("Testing OSC status endpoint...")
    
    # Create client to send to PPT Redirector
    client = udp_client.SimpleUDPClient("127.0.0.1", 9001)
    
    try:
        # Send status request with reply port 9003
        client.send_message("/ppt/status", 9003)
        print("Status request sent to PPT Redirector on port 9001")
        print("Reply should be sent to port 9003")
        print("Check the PPT Redirector console for output")
        
    except Exception as e:
        print(f"Error sending OSC message: {e}")

def test_osc_select():
    """Test the OSC select endpoint"""
    print("\nTesting OSC select endpoint...")
    
    client = udp_client.SimpleUDPClient("127.0.0.1", 9001)
    
    try:
        client.send_message("/ppt/select", "")  # Send empty string as value
        print("Select request sent - window selector should open")
        
    except Exception as e:
        print(f"Error sending OSC select message: {e}")

def test_osc_reset():
    """Test the OSC reset endpoint"""
    print("\nTesting OSC reset endpoint...")
    
    client = udp_client.SimpleUDPClient("127.0.0.1", 9001)
    
    try:
        client.send_message("/ppt/reset", 9003)
        print("Reset request sent - target should be cleared")
        
    except Exception as e:
        print(f"Error sending OSC reset message: {e}")

if __name__ == "__main__":
    print("OSC Test Script for PPT Redirector")
    print("=" * 40)
    
    test_osc_status()
    time.sleep(1)
    
    test_osc_select()
    time.sleep(1)
    
    test_osc_reset()
    
    print("\nTest complete. Check PPT Redirector console for responses.")
