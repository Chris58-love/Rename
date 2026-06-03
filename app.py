from universal_rename_tool.utils import is_colab


def main():
    if is_colab():
        from universal_rename_tool.ui_colab import launch_colab_app
        launch_colab_app()
    else:
        from universal_rename_tool.ui_local import launch_local_app
        launch_local_app()


if __name__ == "__main__":
    main()
