from project import *
import pytest

def main():
    test_AllLoads()
    test_AllGroups()
    test_MaxMoFunc()
    test_MakeTable()

def test_AllLoads():
    assert AllLoads() == (1, 2, 3, 100, 101, 102, 103, 104, 105, 106)
    with pytest.raises(TypeError):
        AllLoads("Hello")
    with pytest.raises(TypeError):
        AllLoads(123)

def test_AllGroups():
    assert AllGroups() == ['_GB1', '_GB2', '_GB3']
    with pytest.raises(TypeError):
        AllGroups("Hello")
    with pytest.raises(TypeError):
        AllGroups(123)

def test_MaxMoFunc():
    assert MaxMoFunc((1, 2, 3, 100, 101, 102, 103, 104, 105, 106), ['_GB1', '_GB2', '_GB3']) == [[15, 2916.65, 101, '_GB1'], [19, 2058.89, 101, '_GB2'], [18, 2330.594, 101, '_GB3']]
    with pytest.raises(TypeError):
        MaxMoFunc(100, 123)
    with pytest.raises(TypeError):
        MaxMoFunc("Hello", 123)
    with pytest.raises(TypeError):
        MaxMoFunc([])

def test_MakeTable():
    with pytest.raises(IndexError):
        MakeTable("Hello")
    with pytest.raises(TypeError):
        MakeTable(101)

if __name__ == "__main__":
    main()
