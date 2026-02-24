def _pick_launcher():
    try:
        from .ui_qt import launch

        return launch
    except ModuleNotFoundError:
        from .ui_tk import launch

        return launch


def main() -> int:
    launch = _pick_launcher()
    launch()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
