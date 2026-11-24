"""Standalone entry point: attach reply reports from a folder, only."""

import argparse
import logging
import os


def main(argv=None):
    parser = argparse.ArgumentParser(
        description='АСУД — прикрепление готовых ответов к отчётам начальника отдела')
    parser.add_argument('--folder', help='Папка ответов; по умолчанию из settings.json')
    parser.add_argument('--watch', action='store_true', help='Постоянно проверять папку')
    parser.add_argument('--headless', action='store_true', help='Запустить Edge в фоне')
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%H:%M:%S',
    )
    if args.folder:
        os.environ['ASUD_REPLY_FOLDER'] = args.folder
    os.environ['ASUD_REPLY_WATCH'] = '1' if args.watch else '0'
    if args.headless:
        os.environ['ASUD_HEADLESS'] = '1'

    from flows.attach_replies import main as attach_reports
    try:
        attach_reports()
    except KeyboardInterrupt:
        logging.getLogger('asud.replies').info('Обработка остановлена.')
    except Exception:
        logging.getLogger('asud.replies').exception('Прикрепление отчётов остановлено из-за ошибки')
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
