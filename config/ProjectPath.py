import os


def get_project_path():
    """window获取项目地址"""
    path = os.path.split(os.path.abspath(__file__))[0]
    project_path = os.path.realpath(os.path.join(path, '..'))
    return project_path
