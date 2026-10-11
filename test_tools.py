from app import calculator, word_counter, is_safe_path

def test_calculator_addition():
    assert calculator("add", 2, 3) == 5

def test_calculator_division():
    assert calculator("divide", 10, 2) == 5

def test_calculator_divide_by_zero():
    result = calculator("divide", 10, 0)
    assert result == "Error: cannot divide by zero"

def test_word_counter():
    assert word_counter("the quick brown fox") == 4

def test_word_counter_empty_string():
    assert word_counter("") == 0

def test_is_safe_path_blocks_outside_files():
    result = is_safe_path("C:\\Windows\\System32\\drivers\\etc\\hosts")
    assert result == False

def test_is_safe_path_allows_project_files():
    result = is_safe_path("employees.csv")
    assert result == True