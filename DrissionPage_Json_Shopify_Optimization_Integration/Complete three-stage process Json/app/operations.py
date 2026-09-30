"""四个阶段的 GUI 适配，所有输入均在启动时快照。"""
from pathlib import Path
from stages.link_collection import collector, config as link_config, logger as link_logger
from stages.product_collection import scraper_runner, browser_utils, file_utils
from stages.data_processing.processing_runner import process_folder


def collect_links(settings):
    """根据已快照的 settings 构造第一阶段闭包（启动浏览器采集分类链接）。"""
    def run(controller, log):
        """执行链接采集，返回输出目录、链接数与完成/部分失败/无数据状态。"""
        cfg = link_config.CollectConfig(
            categories=collector.parse_categories(settings['categories']),
            output_dir=settings['link_output'], max_pages=settings['max_pages'],
            xpath=settings['xpath'], proxy=settings['proxy'] or None)
        runner = collector.CollectorRunner(cfg, controller.stop_event,
            on_progress=lambda i, total, title: controller.emit('progress', (i-1, total, title)),
            on_count=lambda count: controller.emit('count', f'{count} 条链接'))
        controller._interrupt = runner.interrupt
        link_logger.add_sink(log)
        try:
            rows = runner.run()
        finally:
            link_logger.remove_sink(log)
        report = runner.last_report
        if report is None or report.error:
            raise RuntimeError(report.error if report else '未取得保存报告')
        warnings = runner.errors or report.index_error or report.recovery_files or report.json_recovery_files
        return {'output': cfg.output_dir, 'count': len(rows),
                'status': '部分失败' if warnings else ('完成' if rows else '无数据')}
    return run


def collect_products(settings, files):
    """根据已快照的 settings 与文件/图片规则，构造第二阶段采集闭包。"""
    def run(controller, log):
        """执行商品 JSON 采集，返回输出目录、行数及完成/部分失败/无数据状态。"""
        # 浏览器代理从界面输入覆盖；汇率请求走 Python 自身网络代理环境
        browser_utils.PROXY_SERVER = settings['proxy'] or None
        runner = scraper_runner.ScraperRunner(files, settings['product_input'], controller.stop_event,
            on_file_status=lambda i, text, state: controller.emit('file_status', (i, text)),
            on_progress=lambda i, total, text: controller.emit('progress', (i, total, text)))
        rows = runner.run()
        output = str(Path(file_utils.build_output_file_path(files[0]['path'])).parent)
        return {'output': output, 'count': len(rows),
                'status': '部分失败' if runner.failed_count or runner.recovery_count else ('完成' if rows else '无数据')}
    return run


def process_data(settings):
    """根据已快照的 settings 构造第三阶段闭包（字段规范与合并）。"""
    def run(controller, log):
        """执行规范与合并，返回结果目录、合并行数；状态固定为完成。

        页面上勾选参与转换的选项在这里生效：未勾选的维度不会写进生成
        的中间表 styles1，因此行数真正减少。
        """
        report = process_folder(
            settings['processing_input'], log,
            stop_event=controller.stop_event,
            keep_options=settings.get('expand_options'),
        )
        return {'output': str(report.output_dir), 'merged_file': str(report.merged_file), 'count': report.product_rows, 'status': '完成'}
    return run


def convert_data(settings):
    """独立执行第四阶段，不重复清洗或合并输入。"""
    from stages.data_processing.conversion_runner import convert_file
    # None/缺失表示「不限制」，由转换层理解为全部展开。
    expand_options = settings.get('expand_options')
    def run(controller, log):
        return convert_file(
            settings['conversion_input'], log, controller.stop_event,
            expand_options=expand_options,
        )
    return run
