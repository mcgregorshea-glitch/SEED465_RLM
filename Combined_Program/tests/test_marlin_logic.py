import sys
import os
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from sim.marlin_logic import MarlinState

def test_marlin_state_initialization():
    state = MarlinState()
    assert state.x == 0.0
    assert state.y == 0.0
    assert state.z == 0.0
    assert state.e == 0.0
    assert state.absolute_positioning is True

def test_absolute_positioning():
    state = MarlinState()
    state.update_position(x=10, y=20)
    assert state.x == 10.0
    assert state.y == 20.0
    
    state.update_position(x=5)
    assert state.x == 5.0
    assert state.y == 20.0

def test_relative_positioning():
    state = MarlinState()
    state.absolute_positioning = False
    state.update_position(x=10, y=10)
    assert state.x == 10.0
    
    state.update_position(x=5, y=-2)
    assert state.x == 15.0
    assert state.y == 8.0

def test_clamping():
    state = MarlinState()
    state.update_position(x=500, z=-50)
    assert state.x == 220.0  # Max X
    assert state.z == 0.0    # Min Z

def test_homing():
    state = MarlinState()
    state.update_position(x=50, y=50, z=50)
    state.home()
    assert state.x == 0.0
    assert state.y == 0.0
    assert state.z == 0.0

def test_m114_formatting():
    state = MarlinState()
    state.update_position(x=10.5, y=20.75, z=5.0, e=100.0)
    expected = "X:10.50 Y:20.75 Z:5.00 E:100.00 Count X:0 Y:0 Z:0"
    assert state.format_m114() == expected

from sim.marlin_logic import GCodeParser

def test_parser_g1_absolute():
    state = MarlinState()
    parser = GCodeParser(state)
    response = parser.parse_line("G1 X100 Y50 Z10 E5")
    assert response == "ok"
    assert state.x == 100.0
    assert state.y == 50.0
    assert state.z == 10.0
    assert state.e == 5.0

def test_parser_g28_homing():
    state = MarlinState()
    parser = GCodeParser(state)
    state.update_position(x=50, y=50)
    parser.parse_line("G28")
    assert state.x == 0.0
    assert state.y == 0.0

def test_parser_m114():
    state = MarlinState()
    parser = GCodeParser(state)
    state.update_position(x=10, y=20)
    response = parser.parse_line("M114")
    assert "X:10.00 Y:20.00" in response
    assert response.endswith("ok")

def test_parser_relative_mode():
    state = MarlinState()
    parser = GCodeParser(state)
    parser.parse_line("G91")
    parser.parse_line("G1 X10 Y10")
    assert state.x == 10.0
    parser.parse_line("G1 X5")
    assert state.x == 15.0

def test_parser_comments():
    state = MarlinState()
    parser = GCodeParser(state)
    response = parser.parse_line("G1 X10 ; this is a comment")
    assert response == "ok"
    assert state.x == 10.0
    
    response = parser.parse_line("; just a comment")
    assert response == "ok"
