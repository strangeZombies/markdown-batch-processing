#!/usr/bin/env python3
"""
Markdown Frontmatter 专业处理器 - 多语言优化版 v1.4
========================================

功能清单：
1. 解析和序列化 YAML frontmatter
2. 自动检测字段类型并分析冲突（改进 str 到 list 转换）
3. 支持多种数据类型的智能转换
4. 字段合并与默认值填充
5. 交互式变更确认对话框
6. GUI展示检测结果（字段类型与冲突）
7. 生成详细的Excel分析报告（修正输出路径）
8. 支持PyQt6 GUI与命令行模式
9. 配置文件加载与保存（包含语言偏好和转换设置）
10. 多语言支持（联合国通用语言）与动态语言切换

说明：
本工具由 DeepSeek 和 Grok 生成，旨在提供高效的 Markdown frontmatter 处理能力，支持多种语言的用户使用。通过智能分析和转换，用户可以轻松管理和优化其 Markdown 文件的 frontmatter 数据。
"""

import yaml
import os
import sys
from pathlib import Path
from typing import Dict, Tuple, Any, List, Optional, DefaultDict
from collections import defaultdict
from datetime import datetime, date
import pandas as pd
import webbrowser
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QPushButton, QFileDialog, QCheckBox, QTabWidget, QTableWidget,
    QTableWidgetItem, QTextEdit, QMessageBox, QDialog, QDialogButtonBox, QLabel,
    QComboBox, QProgressBar, QHeaderView
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QTextCursor, QTextCharFormat, QColor, QFont

# ====================
# 多语言支持
# ====================
class LanguageManager:
    """管理多语言翻译，支持联合国通用语言"""
    LANGUAGES = {
        'en': {  # 英语
            'window_title': 'Markdown Frontmatter Processor',
            'input_dir': 'Input Directory:',
            'output_dir': 'Output Directory:',
            'overwrite': 'Overwrite Source Files',
            'ignore_null': 'Ignore Null Conflicts',
            'config_file': 'Config File:',
            'language': 'Language:',
            'field_types': 'Field Types',
            'merge_rules': 'Merge Rules',
            'default_values': 'Default Values',
            'field_name': 'Field Name',
            'target_type': 'Target Type',
            'target_field': 'Target Field',
            'source_fields': 'Source Fields',
            'type': 'Type',
            'default_value': 'Default Value',
            'add_field': 'Add Field',
            'add_rule': 'Add Rule',
            'add_default': 'Add Default Rule',
            'analyze': 'Analyze',
            'process': 'Process',
            'view_report': 'View Report',
            'browse': 'Browse...',
            'confirm_changes': 'Confirm Changes - {}',
            'no_valid_files': 'No valid Markdown files found',
            'analysis_complete': 'Analysis complete, report generated: {}',
            'processing_complete': 'Processing complete, report generated: {}',
            'no_report': 'No report file available',
            'error_opening_report': 'Failed to open report: {}',
            'invalid_input_dir': 'Please enter a valid input directory',
            'invalid_output_dir': 'Please specify output directory or select overwrite mode',
            'cannot_create_output': 'Cannot create output directory: {}',
            'config_not_found': 'Config file not found: {}',
            'config_load_failed': 'Failed to load config: {}',
            'config_saved': 'Config saved to: {}',
            'config_save_failed': 'Failed to save config: {}',
            'config_path_missing': 'Please specify config file path',
            'processing_file': 'Processing: {}',
            'file_processed': 'File {} processed, changes: {}',
            'type_conflict': 'Type conflict detected - {}: {}',
            'no_conflicts': 'No type conflicts detected',
            'conflicts_found': 'Found {} fields with type conflicts',
            'invalid_frontmatter': 'File has no valid frontmatter, skipping',
            'yaml_error': 'YAML parsing error: {}'
        },
        'zh': {  # 中文
            'window_title': 'Markdown Frontmatter 专业处理器',
            'input_dir': '输入目录：',
            'output_dir': '输出目录：',
            'overwrite': '直接覆盖源文件',
            'ignore_null': '忽略null值冲突',
            'config_file': '配置文件：',
            'language': '语言：',
            'field_types': '字段类型',
            'merge_rules': '合并规则',
            'default_values': '默认值',
            'field_name': '字段名称',
            'target_type': '目标类型',
            'target_field': '目标字段',
            'source_fields': '合并来源字段',
            'type': '类型',
            'default_value': '默认值',
            'add_field': '添加字段',
            'add_rule': '添加规则',
            'add_default': '添加默认值规则',
            'analyze': '分析检测',
            'process': '开始处理',
            'view_report': '查看报告',
            'browse': '浏览...',
            'confirm_changes': '确认更改 - {}',
            'no_valid_files': '未找到有效的Markdown文件',
            'analysis_complete': '分析完成，报告已生成：{}',
            'processing_complete': '处理完成，报告已生成：{}',
            'no_report': '没有可用的报告文件',
            'error_opening_report': '无法打开报告：{}',
            'invalid_input_dir': '请输入有效的输入目录',
            'invalid_output_dir': '请指定输出目录或选择覆盖模式',
            'cannot_create_output': '无法创建输出目录：{}',
            'config_not_found': '配置文件不存在：{}',
            'config_load_failed': '加载配置文件失败：{}',
            'config_saved': '配置已保存到：{}',
            'config_save_failed': '保存配置失败：{}',
            'config_path_missing': '请先指定配置文件路径',
            'processing_file': '正在处理：{}',
            'file_processed': '文件 {} 已处理，变更：{} 处',
            'type_conflict': '检测到类型冲突 - {}：{}',
            'no_conflicts': '未检测到类型冲突',
            'conflicts_found': '共检测到 {} 个字段存在类型冲突',
            'invalid_frontmatter': '文件无有效frontmatter，跳过',
            'yaml_error': 'YAML解析错误：{}'
        },
        'fr': {  # 法语
            'window_title': 'Processeur Markdown Frontmatter',
            'input_dir': 'Répertoire d\'entrée :',
            'output_dir': 'Répertoire de sortie :',
            'overwrite': 'Écraser les fichiers source',
            'ignore_null': 'Ignorer les conflits de valeurs nulles',
            'config_file': 'Fichier de configuration :',
            'language': 'Langue :',
            'field_types': 'Types de champs',
            'merge_rules': 'Règles de fusion',
            'default_values': 'Valeurs par défaut',
            'field_name': 'Nom du champ',
            'target_type': 'Type cible',
            'target_field': 'Champ cible',
            'source_fields': 'Champs sources à fusionner',
            'type': 'Type',
            'default_value': 'Valeur par défaut',
            'add_field': 'Ajouter un champ',
            'add_rule': 'Ajouter une règle',
            'add_default': 'Ajouter une règle par défaut',
            'analyze': 'Analyser',
            'process': 'Démarrer le traitement',
            'view_report': 'Voir le rapport',
            'browse': 'Parcourir...',
            'confirm_changes': 'Confirmer les modifications - {}',
            'no_valid_files': 'Aucun fichier Markdown valide trouvé',
            'analysis_complete': 'Analyse terminée, rapport généré : {}',
            'processing_complete': 'Traitement terminé, rapport généré : {}',
            'no_report': 'Aucun fichier de rapport disponible',
            'error_opening_report': 'Échec de l\'ouverture du rapport : {}',
            'invalid_input_dir': 'Veuillez entrer un répertoire d\'entrée valide',
            'invalid_output_dir': 'Veuillez spécifier un répertoire de sortie ou sélectionner le mode écrasement',
            'cannot_create_output': 'Impossible de créer le répertoire de sortie : {}',
            'config_not_found': 'Fichier de configuration introuvable : {}',
            'config_load_failed': 'Échec du chargement de la configuration : {}',
            'config_saved': 'Configuration enregistrée sous : {}',
            'config_save_failed': 'Échec de l\'enregistrement de la configuration : {}',
            'config_path_missing': 'Veuillez d\'abord spécifier le chemin du fichier de configuration',
            'processing_file': 'Traitement : {}',
            'file_processed': 'Fichier {} traité, modifications : {}',
            'type_conflict': 'Conflit de type détecté - {} : {}',
            'no_conflicts': 'Aucun conflit de type détecté',
            'conflicts_found': '{} champs avec des conflits de type détectés',
            'invalid_frontmatter': 'Le fichier n\'a pas de frontmatter valide, ignoré',
            'yaml_error': 'Erreur d\'analyse YAML : {}'
        },
        'es': {  # 西班牙语
            'window_title': 'Procesador de Frontmatter Markdown',
            'input_dir': 'Directorio de entrada:',
            'output_dir': 'Directorio de salida:',
            'overwrite': 'Sobrescribir archivos fuente',
            'ignore_null': 'Ignorar conflictos de valores nulos',
            'config_file': 'Archivo de configuración:',
            'language': 'Idioma:',
            'field_types': 'Tipos de campos',
            'merge_rules': 'Reglas de fusión',
            'default_values': 'Valores predeterminados',
            'field_name': 'Nombre del campo',
            'target_type': 'Tipo objetivo',
            'target_field': 'Campo objetivo',
            'source_fields': 'Campos fuente a fusionar',
            'type': 'Tipo',
            'default_value': 'Valor predeterminado',
            'add_field': 'Agregar campo',
            'add_rule': 'Agregar regla',
            'add_default': 'Agregar regla predeterminada',
            'analyze': 'Analizar',
            'process': 'Iniciar procesamiento',
            'view_report': 'Ver informe',
            'browse': 'Examinar...',
            'confirm_changes': 'Confirmar cambios - {}',
            'no_valid_files': 'No se encontraron archivos Markdown válidos',
            'analysis_complete': 'Análisis completado, informe generado: {}',
            'processing_complete': 'Procesamiento completado, informe generado: {}',
            'no_report': 'No hay archivo de informe disponible',
            'error_opening_report': 'Error al abrir el informe: {}',
            'invalid_input_dir': 'Por favor, ingrese un directorio de entrada válido',
            'invalid_output_dir': 'Por favor, especifique el directorio de salida o seleccione el modo de sobrescritura',
            'cannot_create_output': 'No se puede crear el directorio de salida: {}',
            'config_not_found': 'Archivo de configuración no encontrado: {}',
            'config_load_failed': 'Error al cargar la configuración: {}',
            'config_saved': 'Configuración guardada en: {}',
            'config_save_failed': 'Error al guardar la configuración: {}',
            'config_path_missing': 'Por favor, especifique primero la ruta del archivo de configuración',
            'processing_file': 'Procesando: {}',
            'file_processed': 'Archivo {} procesado, cambios: {}',
            'type_conflict': 'Conflicto de tipo detectado - {}: {}',
            'no_conflicts': 'No se detectaron conflictos de tipo',
            'conflicts_found': 'Se encontraron {} campos con conflictos de tipo',
            'invalid_frontmatter': 'El archivo no tiene frontmatter válido, omitido',
            'yaml_error': 'Error de análisis YAML: {}'
        },
        'ar': {  # 阿拉伯语
            'window_title': 'معالج Frontmatter لـ Markdown',
            'input_dir': 'دليل الإدخال:',
            'output_dir': 'دليل الإخراج:',
            'overwrite': 'الكتابة فوق الملفات المصدر',
            'ignore_null': 'تجاهل تعارضات القيم الفارغة',
            'config_file': 'ملف التكوين:',
            'language': 'اللغة:',
            'field_types': 'أنواع الحقول',
            'merge_rules': 'قواعد الدمج',
            'default_values': 'القيم الافتراضية',
            'field_name': 'اسم الحقل',
            'target_type': 'نوع الهدف',
            'target_field': 'الحقل المستهدف',
            'source_fields': 'حقول المصدر للدمج',
            'type': 'النوع',
            'default_value': 'القيمة الافتراضية',
            'add_field': 'إضافة حقل',
            'add_rule': 'إضافة قاعدة',
            'add_default': 'إضافة قاعدة افتراضية',
            'analyze': 'تحليل',
            'process': 'بدء المعالجة',
            'view_report': 'عرض التقرير',
            'browse': 'تصفح...',
            'confirm_changes': 'تأكيد التغييرات - {}',
            'no_valid_files': 'لم يتم العثور على ملفات Markdown صالحة',
            'analysis_complete': 'اكتمل التحليل، تم إنشاء التقرير: {}',
            'processing_complete': 'اكتمل المعالجة، تم إنشاء التقرير: {}',
            'no_report': 'لا يوجد ملف تقرير متاح',
            'error_opening_report': 'فشل في فتح التقرير: {}',
            'invalid_input_dir': 'يرجى إدخال دليل إدخال صالح',
            'invalid_output_dir': 'يرجى تحديد دليل الإخراج أو اختيار وضع الكتابة فوق',
            'cannot_create_output': 'لا يمكن إنشاء دليل الإخراج: {}',
            'config_not_found': 'ملف التكوين غير موجود: {}',
            'config_load_failed': 'فشل في تحميل التكوين: {}',
            'config_saved': 'تم حفظ التكوين في: {}',
            'config_save_failed': 'فشل في حفظ التكوين: {}',
            'config_path_missing': 'يرجى تحديد مسار ملف التكوين أولاً',
            'processing_file': 'جاري المعالجة: {}',
            'file_processed': 'تمت معالجة الملف {}، التغييرات: {}',
            'type_conflict': 'تم اكتشاف تعارض في النوع - {}: {}',
            'no_conflicts': 'لم يتم اكتشاف تعارضات في النوع',
            'conflicts_found': 'تم العثور على {} حقول مع تعارضات في النوع',
            'invalid_frontmatter': 'الملف لا يحتوي على frontmatter صالح، يتم تخطيه',
            'yaml_error': 'خطأ في تحليل YAML: {}'
        },
        'ru': {  # 俄语
            'window_title': 'Обработчик Frontmatter Markdown',
            'input_dir': 'Входная директория:',
            'output_dir': 'Выходная директория:',
            'overwrite': 'Перезаписать исходные файлы',
            'ignore_null': 'Игнорировать конфликты null-значений',
            'config_file': 'Файл конфигурации:',
            'language': 'Язык:',
            'field_types': 'Типы полей',
            'merge_rules': 'Правила слияния',
            'default_values': 'Значения по умолчанию',
            'field_name': 'Имя поля',
            'target_type': 'Целевой тип',
            'target_field': 'Целевое поле',
            'source_fields': 'Поля источника для слияния',
            'type': 'Тип',
            'default_value': 'Значение по умолчанию',
            'add_field': 'Добавить поле',
            'add_rule': 'Добавить правило',
            'add_default': 'Добавить правило по умолчанию',
            'analyze': 'Анализировать',
            'process': 'Начать обработку',
            'view_report': 'Просмотреть отчет',
            'browse': 'Обзор...',
            'confirm_changes': 'Подтвердить изменения - {}',
            'no_valid_files': 'Не найдено действительных файлов Markdown',
            'analysis_complete': 'Анализ завершен, отчет сгенерирован: {}',
            'processing_complete': 'Обработка завершена, отчет сгенерирован: {}',
            'no_report': 'Нет доступного файла отчета',
            'error_opening_report': 'Не удалось открыть отчет: {}',
            'invalid_input_dir': 'Пожалуйста, укажите действительную входную директорию',
            'invalid_output_dir': 'Пожалуйста, укажите выходную директорию или выберите режим перезаписи',
            'cannot_create_output': 'Не удалось создать выходную директорию: {}',
            'config_not_found': 'Файл конфигурации не найден: {}',
            'config_load_failed': 'Не удалось загрузить конфигурацию: {}',
            'config_saved': 'Конфигурация сохранена в: {}',
            'config_save_failed': 'Не удалось сохранить конфигурацию: {}',
            'config_path_missing': 'Пожалуйста, сначала укажите путь к файлу конфигурации',
            'processing_file': 'Обработка: {}',
            'file_processed': 'Файл {} обработан, изменений: {}',
            'type_conflict': 'Обнаружен конфликт типов - {}: {}',
            'no_conflicts': 'Конфликты типов не обнаружены',
            'conflicts_found': 'Обнаружено {} полей с конфликтами типов',
            'invalid_frontmatter': 'Файл не содержит действительного frontmatter, пропущен',
            'yaml_error': 'Ошибка разбора YAML: {}'
        }
    }
    
    TYPE_DISPLAY = {
        'en': {'str': 'String', 'int': 'Integer', 'float': 'Float', 'bool': 'Boolean', 
               'date': 'Date', 'datetime': 'DateTime', 'list': 'List', 'null': 'Null', 'unknown': 'Unknown'},
        'zh': {'str': '字符串', 'int': '整数', 'float': '浮点数', 'bool': '布尔值', 
               'date': '日期', 'datetime': '日期时间', 'list': '列表', 'null': '空值', 'unknown': '未知'},
        'fr': {'str': 'Chaîne', 'int': 'Entier', 'float': 'Flottant', 'bool': 'Booléen', 
               'date': 'Date', 'datetime': 'DateHeure', 'list': 'Liste', 'null': 'Nul', 'unknown': 'Inconnu'},
        'es': {'str': 'Cadena', 'int': 'Entero', 'float': 'Flotante', 'bool': 'Booleano', 
               'date': 'Fecha', 'datetime': 'FechaHora', 'list': 'Lista', 'null': 'Nulo', 'unknown': 'Desconocido'},
        'ar': {'str': 'سلسلة', 'int': 'عدد صحيح', 'float': 'عائم', 'bool': 'منطقي', 
               'date': 'تاريخ', 'datetime': 'تاريخ ووقت', 'list': 'قائمة', 'null': 'فارغ', 'unknown': 'غير معروف'},
        'ru': {'str': 'Строка', 'int': 'Целое', 'float': 'Плавающее', 'bool': 'Булево', 
               'date': 'Дата', 'datetime': 'ДатаВремя', 'list': 'Список', 'null': 'Пусто', 'unknown': 'Неизвестно'}
    }

    def __init__(self, lang: str = 'en'):
        self.lang = lang if lang in self.LANGUAGES else 'en'
    
    def get(self, key: str) -> str:
        """获取翻译文本"""
        return self.LANGUAGES[self.lang].get(key, key)
    
    def get_type_display(self, type_name: str) -> str:
        """获取字段类型的显示名称"""
        return self.TYPE_DISPLAY[self.lang].get(type_name, type_name)

# ====================
# 工具类定义
# ====================

class ChangesDialog(QDialog):
    """交互式变更确认对话框，显示文件变更详情"""
    def __init__(self, file_path: str, changes: List[Dict[str, Any]], lang: LanguageManager, parent=None):
        super().__init__(parent)
        self.lang = lang
        self.setWindowTitle(self.lang.get('confirm_changes').format(Path(file_path).name))
        self.setMinimumSize(600, 400)
        
        layout = QVBoxLayout()
        
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        
        for change in changes:
            self._append_change(
                f"【{change['key']}】 {change.get('action', '修改')}:",
                QColor(0, 0, 255)
            )
            old_val = f"{change['old_value']}" if change['old_value'] is not None else "空"
            new_val = f"{change['new_value']}" if change['new_value'] is not None else "空"
            self._append_change(f"  原值: {old_val}", QColor(128, 0, 0))
            self._append_change(f"  新值: {new_val}", QColor(0, 100, 0))
            self.text_edit.append("")
        
        layout.addWidget(self.text_edit)
        
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def _append_change(self, text: str, color: QColor):
        """添加带颜色格式的变更文本"""
        cursor = self.text_edit.textCursor()
        format_ = QTextCharFormat()
        format_.setForeground(color)
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertText(text + "\n", format_)

class ProcessingThread(QThread):
    """文件处理线程，避免GUI冻结"""
    progress_updated = pyqtSignal(int, str)  # 进度值, 当前文件
    message_logged = pyqtSignal(str, str)    # 消息内容, 类型(info/warn/error)
    processing_finished = pyqtSignal(str)    # 报告路径
    
    def __init__(self, processor, config: Dict[str, Any], lang: LanguageManager):
        super().__init__()
        self.processor = processor
        self.config = config
        self.lang = lang
    
    def run(self):
        """线程主逻辑：批量处理文件并生成报告"""
        try:
            report_path = self.processor.process_directory(
                input_dir=self.config['input_dir'],
                output_dir=self.config['output_dir'],
                merge_map=self.config['merge_map'],
                field_types=self.config['field_types'],
                default_values=self.config['default_values'],
                ignore_null_conflicts=self.config['ignore_null_conflicts'],
                overwrite=self.config['overwrite'],
                thread=self,
                lang=self.lang,
                list_separators=self.config.get('list_separators', [',', ';', '|'])
            )
            self.processing_finished.emit(report_path)
        except Exception as e:
            self.message_logged.emit(f"处理过程中发生致命错误: {str(e)}", "error")

class FrontmatterAnalyzer:
    """Frontmatter分析与处理核心类"""
    
    SUPPORTED_TYPES = {
        'str': lambda v: str(v),
        'int': lambda v: int(float(str(v))) if isinstance(v, (float, str)) else int(v),
        'float': lambda v: float(str(v)) if isinstance(v, (int, str)) else float(v),
        'bool': lambda v: str(v).lower() in ('true', '1', 'yes', 'on'),
        'date': lambda v: FrontmatterAnalyzer._convert_to_date(v),
        'datetime': lambda v: FrontmatterAnalyzer._convert_to_datetime(v),
        'list': lambda v: FrontmatterAnalyzer._convert_to_list(v)
    }

    @staticmethod
    def _convert_to_date(value: Any) -> Optional[date]:
        """尝试将值转换为日期类型"""
        if isinstance(value, date):
            return value
        if isinstance(value, datetime):
            return value.date()
        try:
            return datetime.strptime(str(value), '%Y-%m-%d').date()
        except (ValueError, TypeError):
            try:
                return datetime.fromisoformat(str(value).replace(' ', 'T').replace('Z', '+00:00')).date()
            except (ValueError, TypeError):
                return None
    
    @staticmethod
    def _convert_to_datetime(value: Any) -> Optional[datetime]:
        """尝试将值转换为日期时间类型"""
        if isinstance(value, datetime):
            return value
        if isinstance(value, date) and not isinstance(value, datetime):
            return datetime.combine(value, datetime.min.time())
        try:
            return datetime.fromisoformat(str(value).replace(' ', 'T').replace('Z', '+00:00'))
        except (ValueError, TypeError):
            try:
                return datetime.strptime(str(value), '%Y-%m-%d %H:%M:%S')
            except (ValueError, TypeError):
                return None
    
    @staticmethod
    def _convert_to_list(value: Any, separators: List[str] = [',', ';', '|']) -> List[Any]:
        """将值转换为列表，支持多种分隔符和 YAML 列表格式"""
        if isinstance(value, list):
            return value
        if value is None:
            return []
        value_str = str(value).strip()
        # 检测 YAML 列表格式，例如 "- item" 或 "-"
        if value_str.startswith('-'):
            try:
                parsed = yaml.safe_load(value_str)
                if isinstance(parsed, list):
                    return parsed
            except yaml.YAMLError:
                pass
        # 使用分隔符拆分
        for sep in separators:
            if sep in value_str:
                return [x.strip() for x in value_str.split(sep) if x.strip()]
        # 如果没有分隔符，单值作为列表
        return [value_str] if value_str else []
    
    def __init__(self):
        self.log_callback = None
        self.type_conflicts = defaultdict(lambda: defaultdict(set))  # 字段类型冲突记录
        self.valid_files = []
    
    def log(self, value: str, level: str = "info"):
        """记录日志，调用回调函数或打印到控制台"""
        if isinstance(value, str):
            message = value
        else:
            message = str(value)
        if self.log_callback:
            self.log_callback(message, level)
        else:
            print(f"[{level.upper()}] {message}")
    
    def detect_type(self, value: Any, list_separators: List[str] = None) -> str:
        """检测字段值的类型，改进对列表的识别"""
        if value is None:
            return 'null'
        
        if isinstance(value, list):
            return 'list'
        
        if isinstance(value, (int, float, bool, str, date, datetime)):
            if isinstance(value, int) and not isinstance(value, bool):
                return 'int'
            if isinstance(value, float):
                return 'float'
            if isinstance(value, bool):
                return 'bool'
            if isinstance(value, str):
                # 检查是否为 YAML 列表或包含分隔符
                if value.strip().startswith('-') or any(sep in value for sep in list_separators or [',', ';', '|']):
                    try:
                        parsed = self._convert_to_list(value, list_separators or [',', ';', '|'])
                        if len(parsed) > 0:
                            return 'list'
                    except Exception:
                        pass
                return 'str'
            if isinstance(value, date) and not isinstance(value, datetime):
                return 'date'
            if isinstance(value, datetime):
                return 'datetime'
        
        for type_name, converter in self.SUPPORTED_TYPES.items():
            try:
                result = converter(value, list_separators or [',', ';', '|'])
                if result is not None:
                    return type_name
            except (ValueError, TypeError):
                continue
        return 'unknown'

    def parse_frontmatter(self, content: str) -> Tuple[Optional[Dict[str, Any]], str]:
        """解析 Markdown 文件中的 YAML frontmatter"""
        content = content.strip()
        if not content.startswith('---'):
            self.log('invalid_frontmatter', "warning")
            return None, content
        
        parts = content.split('---', 2)
        if len(parts) < 3:
            self.log('invalid_frontmatter', "warning")
            return None, content
        
        try:
            frontmatter = yaml.safe_load(parts[1])
            if not isinstance(frontmatter, dict):
                self.log('invalid_frontmatter', "warning")
                return None, parts[2].lstrip()
            return frontmatter, parts[2].lstrip()
        except yaml.YAMLError as e:
            self.log(f"yaml_error: {str(e)}", "error")
            return None, content
    
    def analyze_files(self, input_dir: str, lang: LanguageManager, list_separators: List[str]):
        """分析目录中所有 Markdown 文件的 frontmatter"""
        self.type_conflicts.clear()
        self.valid_files.clear()
        
        input_path = Path(input_dir)
        if not input_path.is_dir():
            self.log(f"{lang.get('invalid_input_dir')} {input_dir}", "error")
            return
        
        for filepath in sorted(input_path.rglob("*.[mM][dD]")):
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                frontmatter, _ = self.parse_frontmatter(content)
                if not frontmatter:
                    continue
                
                self.valid_files.append(str(filepath))
                
                for key, value in frontmatter.items():
                    detected_type = self.detect_type(value, list_separators)
                    self.type_conflicts[key][detected_type].add(str(filepath))
                    
            except Exception as e:
                self.log(f"处理文件 {filepath.name} 失败: {str(e)}", "error")

    def process_file(self, filepath: str, output_dir: str, merge_map: Dict[str, List[str]],
                    field_types: Dict[str, str], default_values: Dict[str, Tuple[str, Any]],
                    ignore_null_conflicts: bool, overwrite: bool, thread: Optional[QThread],
                    lang: LanguageManager, list_separators: List[str]) -> List[Dict[str, Any]]:
        """处理单个 Markdown 文件，应用类型转换、合并和默认值"""
        changes = []
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            
            frontmatter, body = self.parse_frontmatter(content)
            if not frontmatter:
                self.log(f"{lang.get('invalid_frontmatter')}: {Path(filepath).name}", "warning")
                return changes
            
            new_frontmatter = frontmatter.copy()
            
            # 1. 应用合并规则
            for target, sources in merge_map.items():
                values = []
                for src in sources:
                    if src in new_frontmatter and new_frontmatter[src] is not None:
                        val = new_frontmatter[src]
                        if isinstance(val, list):
                            values.extend(val)
                        else:
                            values.append(val)
                if values:
                    new_frontmatter[target] = values
                    changes.append({
                        'key': target, 'action': '合并',
                        'old_value': frontmatter.get(target), 'new_value': values
                    })
                for src in sources:
                    if src in new_frontmatter and src != target:
                        del new_frontmatter[src]
                        changes.append({
                            'key': src, 'action': '删除',
                            'old_value': frontmatter.get(src), 'new_value': None
                        })
            
            # 2. 应用类型转换
            for key, target_type in field_types.items():
                if key in new_frontmatter and new_frontmatter[key] is not None:
                    old_val = new_frontmatter[key]
                    current_type = self.detect_type(old_val, list_separators)
                    if current_type != target_type and (not ignore_null_conflicts or current_type != 'null'):
                        try:
                            new_frontmatter[key] = self.SUPPORTED_TYPES[target_type](old_val, list_separators)
                            changes.append({
                                'key': key, 'action': '类型转换',
                                'old_value': old_val, 'new_value': new_frontmatter[key]
                            })
                        except (ValueError, TypeError) as e:
                            self.log(f"文件 {Path(filepath).name} 字段 {key} 类型转换失败: {str(e)}", "warning")
            
            # 3. 应用默认值
            for key, (val_type, default_val) in default_values.items():
                if key not in new_frontmatter or new_frontmatter[key] is None:
                    new_frontmatter[key] = default_val
                    changes.append({
                        'key': key, 'action': '填充默认值',
                        'old_value': None, 'new_value': default_val
                    })
            
            # 4. 保存修改后的文件
            if changes and not overwrite:
                output_file = Path(output_dir) / Path(filepath).relative_to(Path(input_dir))
                output_file.parent.mkdir(parents=True, exist_ok=True)
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write('---\n')
                    yaml.dump(new_frontmatter, f, allow_unicode=True, sort_keys=False)
                    f.write('---\n')
                    f.write(body)
            elif changes and overwrite:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write('---\n')
                    yaml.dump(new_frontmatter, f, allow_unicode=True, sort_keys=False)
                    f.write('---\n')
                    f.write(body)
            
            if thread:
                thread.progress_updated.emit(1, Path(filepath).name)
            
        except Exception as e:
            self.log(f"处理文件 {Path(filepath).name} 失败: {str(e)}", "error")
        
        return changes
    
    def process_directory(self, input_dir: str, output_dir: str, merge_map: Dict[str, List[str]], 
                     field_types: Dict[str, str], default_values: Dict[str, Tuple[str, Any]], 
                     ignore_null_conflicts: bool, overwrite: bool, 
                     thread: Optional[QThread], lang: LanguageManager, 
                     list_separators: List[str]) -> str:
        """批量处理目录中的 Markdown 文件"""
        self.analyze_files(input_dir, lang, list_separators)
        if not self.valid_files:
            self.log(lang.get('no_valid_files'), "warning")
            return ""
        
        processed_files = 0
        for filepath in self.valid_files:
            changes = self.process_file(
                filepath, output_dir, merge_map, field_types, default_values,
                ignore_null_conflicts, overwrite, thread, lang, list_separators
            )
            if changes:
                processed_files += 1
                self.log(f"{lang.get('file_processed').format(Path(filepath).name, len(changes))}", "info")
            if thread:
                thread.progress_updated.emit(processed_files, Path(filepath).name)
        
        report_path = self.generate_report(output_dir if not overwrite else input_dir)
        self.log(f"{lang.get('processing_complete').format(report_path)}", "info")
        return report_path
    
    def generate_report(self, report_dir: str) -> str:
        """生成 Excel 格式的分析报告，修正输出路径"""
        report_path = Path(report_dir) / 'frontmatter_analysis_report.xlsx'
        report_path.parent.mkdir(parents=True, exist_ok=True)
        
        with pd.ExcelWriter(report_path, engine='xlsxwriter') as writer:
            valid_files_df = pd.DataFrame({
                "File Path": [str(Path(f).relative_to(report_path.parent)) for f in self.valid_files]
            })
            valid_files_df.to_excel(writer, sheet_name="Valid Files", index=False)
            
            conflict_data = []
            for field, type_info in self.type_conflicts.items():
                non_null_types = [t for t, f in type_info.items() if t != 'null' and f]
                if len(non_null_types) > 1:
                    for type_name, files in type_info.items():
                        if not files:
                            continue
                        for file in files:
                            conflict_data.append({
                                "Field": field,
                                "Type": type_name,
                                "File": str(Path(file).relative_to(report_path.parent))
                            })
            if conflict_data:
                conflict_df = pd.DataFrame(conflict_data)
                conflict_df.to_excel(writer, sheet_name="Type Conflicts", index=False)
            
            stats_data = []
            for field, type_info in self.type_conflicts.items():
                stats_data.append({
                    "Field": field,
                    "Detected Types": ", ".join(sorted([t for t, f in type_info.items() if f])),
                    "File Count": sum(len(files) for files in type_info.values())
                })
            stats_df = pd.DataFrame(stats_data)
            stats_df.to_excel(writer, sheet_name="Field Statistics", index=False)
        
        return str(report_path)

# ====================
# 主界面类
# ====================

class MainWindow(QMainWindow):
    def __init__(self, lang: LanguageManager = None):
        super().__init__()
        self.lang = lang if lang else LanguageManager('zh')  # 默认中文
        self.setWindowTitle(self.lang.get('window_title'))
        self.setMinimumSize(1000, 800)
        
        self.analyzer = FrontmatterAnalyzer()
        self.analyzer.log_callback = self.log_message
        self.current_report = None
        self.processing_thread = None
        self.list_separators = [',', ';', '|']  # 默认列表分隔符
        
        self.init_ui()
    
    def init_ui(self):
        """初始化图形用户界面"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 文件路径配置区域
        path_group = QWidget()
        path_layout = QFormLayout(path_group)
        
        self.input_dir_edit = QLineEdit()
        input_browse_btn = QPushButton(self.lang.get('browse'))
        input_browse_btn.clicked.connect(self.browse_input_dir)
        input_layout = QHBoxLayout()
        input_layout.addWidget(self.input_dir_edit)
        input_layout.addWidget(input_browse_btn)
        path_layout.addRow(self.lang.get('input_dir'), input_layout)
        
        self.output_dir_edit = QLineEdit()
        output_browse_btn = QPushButton(self.lang.get('browse'))
        output_browse_btn.clicked.connect(self.browse_output_dir)
        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_dir_edit)
        output_layout.addWidget(output_browse_btn)
        path_layout.addRow(self.lang.get('output_dir'), output_layout)
        
        # 语言与选项组
        options_group = QWidget()
        options_layout = QHBoxLayout(options_group)
        
        self.language_combo = QComboBox()
        self.language_combo.addItems(['English', '中文', 'Français', 'Español', 'العربية', 'Русский'])
        self.language_combo.setCurrentText('中文')
        self.language_combo.currentIndexChanged.connect(self.change_language)
        options_layout.addWidget(QLabel(self.lang.get('language')))
        options_layout.addWidget(self.language_combo)
        
        self.overwrite_check = QCheckBox(self.lang.get('overwrite'))
        self.overwrite_check.stateChanged.connect(self.toggle_output_dir)
        options_layout.addWidget(self.overwrite_check)
        
        self.ignore_null_check = QCheckBox(self.lang.get('ignore_null'))
        options_layout.addWidget(self.ignore_null_check)
        
        # 配置选项区域
        config_group = QWidget()
        config_layout = QFormLayout(config_group)
        
        self.config_edit = QLineEdit("frontmatter_config.yaml")
        config_browse_btn = QPushButton(self.lang.get('browse'))
        config_browse_btn.clicked.connect(self.browse_config)
        config_btn_layout = QHBoxLayout()
        config_btn_layout.addWidget(self.config_edit)
        config_btn_layout.addWidget(config_browse_btn)
        config_layout.addRow(self.lang.get('config_file'), config_btn_layout)
        
        # 主选项卡
        self.tab_widget = QTabWidget()
        
        self.field_type_tab = QWidget()
        self.init_field_type_tab()
        self.tab_widget.addTab(self.field_type_tab, self.lang.get('field_types'))
        
        self.merge_tab = QWidget()
        self.init_merge_tab()
        self.tab_widget.addTab(self.merge_tab, self.lang.get('merge_rules'))
        
        self.defaults_tab = QWidget()
        self.init_defaults_tab()
        self.tab_widget.addTab(self.defaults_tab, self.lang.get('default_values'))
        
        # 检测结果区域
        self.results_table = QTableWidget(0, 3)
        self.results_table.setHorizontalHeaderLabels([
            self.lang.get('field_name'), self.lang.get('type'), "Files"
        ])
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        # 日志区域
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        # 按钮区域
        button_group = QWidget()
        button_layout = QHBoxLayout(button_group)
        
        analyze_btn = QPushButton(self.lang.get('analyze'))
        analyze_btn.clicked.connect(self.run_analysis)
        button_layout.addWidget(analyze_btn)
        
        process_btn = QPushButton(self.lang.get('process'))
        process_btn.clicked.connect(self.start_processing)
        button_layout.addWidget(process_btn)
        
        report_btn = QPushButton(self.lang.get('view_report'))
        report_btn.clicked.connect(self.open_report)
        button_layout.addWidget(report_btn)
        
        main_layout.addWidget(path_group)
        main_layout.addWidget(options_group)
        main_layout.addWidget(config_group)
        main_layout.addWidget(self.tab_widget)
        main_layout.addWidget(self.results_table)
        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.log_area)
        main_layout.addWidget(button_group)
    
    def init_field_type_tab(self):
        """初始化字段类型配置选项卡"""
        layout = QVBoxLayout(self.field_type_tab)
        
        self.field_table = QTableWidget(0, 2)
        self.field_table.setHorizontalHeaderLabels([self.lang.get('field_name'), self.lang.get('target_type')])
        self.field_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        add_row_btn = QPushButton(self.lang.get('add_field'))
        add_row_btn.clicked.connect(self.add_field_row)
        
        layout.addWidget(self.field_table)
        layout.addWidget(add_row_btn)
    
    def init_merge_tab(self):
        """初始化合并规则选项卡"""
        layout = QVBoxLayout(self.merge_tab)
        
        self.merge_table = QTableWidget(0, 2)
        self.merge_table.setHorizontalHeaderLabels([self.lang.get('target_field'), self.lang.get('source_fields')])
        self.merge_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        add_row_btn = QPushButton(self.lang.get('add_rule'))
        add_row_btn.clicked.connect(self.add_merge_row)
        
        layout.addWidget(self.merge_table)
        layout.addWidget(add_row_btn)
    
    def init_defaults_tab(self):
        """初始化默认值选项卡"""
        layout = QVBoxLayout(self.defaults_tab)
        
        self.defaults_table = QTableWidget(0, 3)
        self.defaults_table.setHorizontalHeaderLabels([
            self.lang.get('field_name'), self.lang.get('type'), self.lang.get('default_value')
        ])
        self.defaults_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        add_row_btn = QPushButton(self.lang.get('add_default'))
        add_row_btn.clicked.connect(self.add_default_row)
        
        layout.addWidget(self.defaults_table)
        layout.addWidget(add_row_btn)
    
    def add_field_row(self):
        """添加字段类型配置行"""
        row = self.field_table.rowCount()
        self.field_table.insertRow(row)
        
        type_combo = QComboBox()
        type_combo.addItems([
            self.lang.get_type_display('str'), self.lang.get_type_display('int'),
            self.lang.get_type_display('float'), self.lang.get_type_display('bool'),
            self.lang.get_type_display('date'), self.lang.get_type_display('datetime'),
            self.lang.get_type_display('list')
        ])
        type_combo.setItemData(0, 'str', Qt.ItemDataRole.UserRole)
        type_combo.setItemData(1, 'int', Qt.ItemDataRole.UserRole)
        type_combo.setItemData(2, 'float', Qt.ItemDataRole.UserRole)
        type_combo.setItemData(3, 'bool', Qt.ItemDataRole.UserRole)
        type_combo.setItemData(4, 'date', Qt.ItemDataRole.UserRole)
        type_combo.setItemData(5, 'datetime', Qt.ItemDataRole.UserRole)
        type_combo.setItemData(6, 'list', Qt.ItemDataRole.UserRole)
        self.field_table.setCellWidget(row, 1, type_combo)
    
    def add_merge_row(self):
        """添加合并规则行"""
        row = self.merge_table.rowCount()
        self.merge_table.insertRow(row)
    
    def add_default_row(self):
        """添加默认值规则行"""
        row = self.defaults_table.rowCount()
        self.defaults_table.insertRow(row)
        
        type_combo = QComboBox()
        type_combo = ['add_items']
        for type_name in ['str', 'int', 'float', 'bool', 'date', 'datetime', 'list']:
            type_combo.addItem(self.lang.get_type_display(type_name))
            type_combo.setItemData(type_combo.count() - 1, type_name, Qt.ItemDataRole.UserRole)
        self.defaults_table.setCellWidget(row, 1, type_combo)
    
    def browse_input_dir(self):
        """浏览并选择输入目录"""
        dir_path = QFileDialog.getExistingDirectory(self, self.lang.get('input_dir'))
        if dir_path:
            self.input_dir_edit.setText(dir_path)
            if not self.output_dir_edit.text() and not self.overwrite_check.isChecked():
                self.output_dir_edit.setText(str(Path(dir_path) / 'output'))
    
    def browse_output_dir(self):
        """浏览并选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, self.lang.get('output_dir'))
        if dir_path:
            self.output_dir_edit.setText(dir_path)
    
    def browse_config(self):
        """浏览并选择配置文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, self.lang.get('config_file'), filter="YAML Files (*.yaml *.yml)")
        
        if file_path:
            self.config_edit.setText(file_path)
            self.load_config()
    
    def toggle_output_dir(self):
        """切换覆盖模式时控制输出目录输入框启用状态"""
        self.output_dir_edit.setEnabled(not self.overwrite_check.isChecked())
    
    def change_language(self, index):
        """动态切换语言"""
        lang_map = {
            'English': 'en'
            , '中文': 'zh'
            , 'Français': 'fr'
            , 'es':'es'
            , 'العربية': 'ar'
            , 'Русский': 'ru'
        }
        new_lang = lang_map[self.language_combo.itemText(index)]
        self.lang = LanguageManager(new_lang)
        self.update_ui_texts()
    
    def update_ui_texts(self):
        """更新界面文本以反映当前语言"""
        self.setWindowTitle(self.lang.get('window_title'))
        
        # 更新路径标签
        path_group = self.centralWidget().layout().itemAt(0).widget()
        path_layout = path_group.layout()
        path_layout.itemAt(0).labelItem().setText(self.lang.get('input_dir'))
        path_layout.itemAt(1).layout().itemAt(1).widget().setText(self.lang.get('browse'))
        path_layout.itemAt(1).labelItem().setText(self.lang.get('output_dir'))
        path_layout.itemAt(1).layout().itemAt(1).widget().setText(self.lang.get('browse'))
        
        # 更新选项组
        options_group = self.centralWidget().layout().itemAt(1).widget()
        options_layout = options_group.layout()
        options_layout.itemAt(0).widget().setText(self.lang.get('language'))
        self.overwrite_check.setText(self.lang.get('overwrite'))
        self.ignore_null_check.setText(self.lang.get('ignore_null'))
        
        # 更新配置文件区域
        config_group = self.centralWidget().layout().itemAt(2).widget()
        config_layout = config_group.layout()
        config_layout.itemAt(0).labelItem().setText(self.lang.get('config_file'))
        config_layout.itemAt(layout(0)).itemAt(1).widget().setText(self.lang.get('browse'))
        
        # 更新选项卡
        self.tab_widget.setTabText(0, self.lang.get('field_types'))
        self.tab_widget.setTabText(1, self.lang.get('merge_rules'))
        self.tab_widget.setTabText(2, self.lang.get('default_values'))
        
        self.field_table.setHorizontalHeaderLabels([self.lang.get("field_name"), self.lang.get("target_type")])
        self.merge_table.setHorizontalHeaderLabels([self.lang.get("target_field"), self.lang.get("source_fields")])
        self.defaults_table.setHorizontalHeaderLabels([
            self.lang.get('field_name'), self.lang.get('type'), self.lang.get('default_value')
        ])
        self.results_table.setHorizontalHeaderLabels([
            self.lang.get('field_name'), self.lang.get('type'), 'Files'
        ])
        
        # 更新按钮
        button_group = self.centralWidget().layout().itemAt(6).widget()
        button_layout = button_group.layout()
        button_layout.itemAt(0).widget().setText(self.lang.get('analyze'))
        button_layout.itemAt(1).widget().setText(self.lang.get('process_button'))
        button_layout.itemAt(2).widget().setText(self.lang.get('view_report'))
    
    def load_config(self):
        """从 YAML 文件加载配置"""
        config_path = self.config_edit.text()
        if not config_path or not os.path.exists(config_path):
            self.log_message(self.lang.get('config_not_found').format(config_path), "warning")
            return
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f) or {}
            
            # 加载语言设置
            if 'language' in config:
                lang_map = {
                    'en': 'English', 'zh': '中文', 'fr': 'Français',
                    'es': 'Español', 'ar': 'العربية', 'ru': 'Русский'
                }
                if config['language'] in lang_map:
                    self.language_combo.setCurrentText(lang_map[config['language']])
                    self.change_language(self.language_combo.currentIndex())
            
            # 加载列表分隔符
            if 'list_separators' in config:
                self.list_separators = config['list_separators']
            
            # 加载字段类型
            self.field_table.setRowCount(0)
            for field, field_type in config.get('field_types', {}).items():
                row = self.field_table.rowCount()
                self.field_table.insertRow(row)
                self.field_table.setRowItem(row, 0, QTableWidgetItem(field))
                combo = QComboBox()
                combo.addItems([
                    self.lang.get_type_display('str'), self.lang.get_type_display('int'),
                    self.lang.get_type_display('float'), self.lang.get_type_display('bool'),
                    self.lang.get_type_display('date'), self.lang.get_type_display('datetime'),
                    self.lang.get_type_display('list')
                ])
                combo.setItemData(0, 'str', Qt.ItemDataRole.UserRole)
                combo.setItemData(1, 'int', Qt.ItemDataRole.UserRole)
                combo.setItemData(2, 'float', Qt.ItemDataRole.UserRole)
                combo.setItemData(3, 'bool', Qt.ItemDataRole.UserRole)
                combo.setItemData(4, 'date', Qt.ItemDataRole.UserRole)
                combo.setItemData(5, 'datetime', Qt.ItemDataRole.UserRole)
                combo.setItemData(6, 'list', Qt.ItemDataRole.UserRole)
                if field_type in FrontmatterAnalyzer.SUPPORTED_TYPES:
                    combo.setCurrentText(self.lang.get_type_display(field_type))
                self.field_table.setCellWidget(row, 1, combo)
            
            # 加载合并规则
            self.merge_table.setRowCount(0)
            for target, sources in config.get('merge_rules', {}).items():
                row = self.merge_table.rowCount()
                self.merge_table.insertRow(row)
                self.merge_table.setItem(row, 0, QTableWidgetItem(target))
                self.merge_table.setItem(row, 1, QTableWidgetItem(",".join(sources)))
            
            # 加载默认值
            self.defaults_table.setRowCount(0)
            for field, value_info in config.get('default_values', {}).items():
                row = self.defaults_table.rowCount()
                self.defaults_table.insertRow(row)
                self.defaults_table.setItem(row, 0, QTableWidgetItem(field))
                
                combo = QComboBox()
                combo.addItems([
                    self.lang.get_type_display('str'), self.lang.get_type_display('int'),
                    self.lang.get_type_display('float'), self.lang.get_type_display('bool'),
                    self.lang.get_type_display('date'), self.lang.get_type_display('datetime'),
                    self.lang.get_type_display('list')
                ])
                combo.setItemData(0, 'str', Qt.ItemDataRole.UserRole)
                combo.setItemData(1, 'int', Qt.ItemDataRole.UserRole)
                combo.setItemData(2, 'float', Qt.ItemDataRole.UserRole)
                combo.setItemData(3, 'bool', Qt.ItemDataRole.UserRole)
                combo.setItemData(4, 'date', Qt.ItemDataRole.UserRole)
                combo.setItemData(5, 'datetime', Qt.ItemDataRole.UserRole)
                combo.setItemData(6, 'list', Qt.ItemDataRole.UserRole)
                
                value_type = value_info.get('type', 'str')
                combo.setCurrentText(self.lang.get_type_display(value_type))
                self.defaults_table.setCellWidget(row, 1, combo)
                self.defaults_table.setItem(row, 2, QTableWidgetItem(str(value_info.get('value', ''))))
            
            self.log_message(self.lang.get('config_saved').format(config_path), "info")
        except Exception as e:
            self.log_message(self.lang.get('config_load_failed').format(str(e)), "error")
    
    def save_config(self):
        """保存当前配置到 YAML 文件，包括语言和分隔符设置"""
        config_path = self.config_edit.text()
        if not config_path:
            self.log_message(self.lang.get('config_path_missing'), "warning")
            return False
        
        config = {
            'language': self.lang.lang,
            'list_separators': self.list_separators,
            'field_types': {},
            'merge_rules': {},
            'default_values': {}
        }
        
        # 保存字段类型
        for row in range(self.field_table.rowCount()):
            field_item = self.field_table.item(row, 0)
            combo = self.field_table.cellWidget(row, 1)
            if field_item and field_item.text() and combo:
                config['field_types'][field_item.text()] = combo.currentData(Qt.ItemDataRole.UserRole)
        
        # 保存合并规则
        for row in range(self.merge_table.rowCount()):
            target_item = self.merge_table.item(row, 0)
            sources_item = self.merge_table.item(row, 1)
            if target_item and target_item.text() and sources_item and sources_item.text():
                sources = [s.strip() for s in sources_item.text().split(",") if s.strip()]
                if sources:
                    config['merge_rules'][target_item.text()] = sources
        
        # 保存默认值
        for row in range(self.defaults_table.rowCount()):
            field_item = self.defaults_table.item(row, 0)
            combo = self.defaults_table.cellWidget(row, 1)
            value_item = self.defaults_table.item(row, 2)
            if field_item and field_item.text() and combo and value_item and value_item.text():
                value_type = combo.currentData(Qt.ItemDataRole.UserRole)
                config['default_values'][field_item.text()] = {
                    'type': value_type,
                    'value': self._convert_config_value(value_item.text(), value_type)
                }
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config, f, allow_unicode=True, sort_keys=False)
            self.log_message(self.lang.get('config_saved').format(config_path), "info")
            return True
        except Exception as e:
            self.log_message(self.lang.get('config_save_failed').format(str(e)), "error")
            return False
    
    def _convert_config_value(self, value: str, value_type: str) -> Any:
        """将配置中的字符串值转换为对应类型"""
        try:
            if value_type == 'int':
                return int(value)
            elif value_type == 'float':
                return float(value)
            elif value_type == 'bool':
                return value.lower() in ('true', '1', 'yes', 'on')
            elif value_type == 'date':
                return datetime.strptime(value, '%Y-%m-%d').date()
            elif value_type == 'list':
                return [v.strip() for v in value.split(",") if v.strip()]
            return value
        except (ValueError, TypeError) as e:
            self.log_message(f"配置值转换失败 ({value_type}): {value} -> {str(e)}", "warning")
            return value
    
    def validate_inputs(self) -> bool:
        """验证用户输入的有效性"""
        input_dir = self.input_dir_edit.text()
        if not input_dir or not os.path.isdir(input_dir):
            self.log_message(self.lang.get('invalid_input_dir'), "error")
            return False
        
        if not self.overwrite_check.isChecked():
            output_dir = self.output_dir_edit.text()
            if not output_dir:
                self.log_message(self.lang.get('invalid_output_dir'), "error")
                return False
            try:
                Path(output_dir).mkdir(parents=True, exist_ok=True)
            except Exception as e:
                self.log_message(self.lang.get('cannot_create_output').format(str(e)), "error")
                return False
        
        return True
    
    def run_analysis(self):
        """执行 frontmatter 分析并在 GUI 中显示结果"""
        if not self.validate_inputs():
            return
        
        self.log_area.clear()
        self.results_table.setRowCount(0)
        input_dir = self.input_dir_edit.text()
        self.log_message(f"{self.lang.get('analyze')} {input_dir}", "info")
        
        self.analyzer.analyze_files(input_dir, self.lang, self.list_separators)
        
        # 显示检测结果
        conflict_count = 0
        for field, type_info in self.analyzer.type_conflicts.items():
            non_null_types = [t for t, files in type_info.items() if t != 'null' and files]
            row = self.results_table.rowCount()
            self.results_table.insertRow(row)
            self.results_table.setItem(row, 0, QTableWidgetItem(field))
            types = [self.lang.get_type_display(t) for t, files in type_info.items() if files]
            self.results_table.setItem(row, 1, QTableWidgetItem(", ".join(sorted(types))))
            self.results_table.setItem(row, 2, QTableWidgetItem(str(sum(len(f) for f in type_info.values()))))
            if len(non_null_types) > 1:
                self.log_message(self.lang.get('type_conflict').format(field, ', '.join(non_null_types)), "warning")
                conflict_count += 1
        
        if conflict_count == 0:
            self.log_message(self.lang.get('no_conflicts'), "info")
        else:
            self.log_message(self.lang.get('conflicts_found').format(conflict_count), "warning")
        
        if self.analyzer.valid_files:
            output_dir = input_dir if self.overwrite_check.isChecked() else self.output_dir_edit.text()
            self.current_report = self.analyzer.generate_report(output_dir)
            self.log_message(self.lang.get('analysis_complete').format(self.current_report), "info")
    
    def start_processing(self):
        """开始批量处理文件"""
        if not self.validate_inputs():
            return
        
        if not self.save_config():
            return
        
        config = {
            'input_dir': self.input_dir_edit.text(),
            'output_dir': self.output_dir_edit.text(),
            'overwrite': self.overwrite_check.isChecked(),
            'ignore_null_conflicts': self.ignore_null_check.isChecked(),
            'list_separators': self.list_separators,
            'field_types': {},
            'merge_map': {},
            'default_values': {}
        }
        
        # 收集字段类型
        for row in range(self.field_table.rowCount()):
            field_item = self.field_table.item(row, 0)
            combo = self.field_table.cellWidget(row, 1)
            if field_item and field_item.text() and combo:
                config['field_types'][field_item.text()] = combo.currentData(Qt.ItemDataRole.UserRole)
        
        # 收集合并规则
        for row in range(self.merge_table.rowCount()):
            target_item = self.merge_table.item(row, 0)
            sources_item = self.merge_table.item(row, 1)
            if target_item and target_item.text() and sources_item and sources_item.text():
                sources = [s.strip() for s in sources_item.text().split(',') if s.strip()]
                if sources:
                    config['merge_map'][target_item.text()] = sources
        
        # 收集默认值配置
        for row in range(self.defaults_table.rowCount()):
            field_item = self.defaults_table.item(row, 0)
            combo = self.defaults_table.cellWidget(row, 1)
            value_item = self.defaults_table.item(row, 2)
            if field_item and field_item.text() and combo and value_item and value_item.text():
                value_type = combo.currentData(Qt.ItemDataRole.UserRole)
                config['default_values'][field_item.text()] = (
                    value_type,
                    self._convert_config_value(value_item.text(), value_type)
                )
        
        # 启动处理线程
        self.log_area.clear()
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        total_files = sum(1 for _ in Path(config['input_dir']).rglob("*.md")) + \
            sum(1 for _ in Path(config['input_dir']).rglob("*.markdown"))
        self.progress_bar.setMaximum(max(total_files, 1))
        
        self.processing_thread = ProcessingThread(self.analyzer, config, self.lang)
        self.processing_thread.progress_updated.connect(self.update_progress)
        self.processing_thread.message_logged.connect(self.log_message)
        self.processing_thread.processing_finished.connect(self.on_processing_finished)
        self.processing_thread.start()
        
    def update_progress(self, value: int, file_name: str):
        """更新进度条和当前处理文件信息"""
        self.progress_bar.setValue(value)
        self.log_message(self.lang.get('processing_file').format(file_name), "info")
        
    def on_processing_finished(self, report_path: str):
        """处理完成后的回调"""
        self.progress_bar.setVisible(False)
        self.current_report = report_path
        self.log_message(self.lang.get('processing_complete').format(report_path), "info")
        QMessageBox.information(self, "完成", self.lang.get('processing_complete').format(report_path))
        self.processing_thread = None
    
    def open_report(self):
        """打开生成的 Excel报告"""
        if not self.current_report or not os.path.exists(self.current_report):
            QMessageBox.warning(self, "警告", self.lang.get('no_report'))
            return
        
        try:
            webbrowser.open(f"file://{self.current_report}")
        except Exception as e:
            QMessageBox.critical(self, "错误", self.lang.get('error_opening_report').format(str(e)))
    
    def log_message(self, message: str, level: str = "info"):
        """格式化并显示日志消息"""
        cursor = self.log_area.textCursor()
        format_ = QTextCharFormat()
        
        if level == "error":
            format_.setForeground(QColor(220, 20, 60))
        elif level == "warning":
            format_.setForeground(QColor(255, 165, 0))
        else:
            format_.setForeground(QColor(0, 0, 0))
        
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.insertText(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ", format_)
        
        format_.setFontWeight(QFont.Weight.Bold if level in ("error", "warning") else QFont.Weight.Normal)
        cursor.insertText(f"[{level.upper()}] ", format_)
        
        format_.setFontWeight(QFont.Weight.Normal)
        cursor.insertText(f"{message}\n", format_)
        
        self.log_area.ensureCursorVisible()
        
        lines = self.log_area.toPlainText().split("\n")
        if len(lines) > 1000:
            self.log_area.setPlainText("\n".join(lines[-1000:]))
    
    def closeEvent(self, event):
        """处理窗口关闭事件"""
        if self.processing_thread and self.processing_thread.isRunning():
            reply = QMessageBox.question(
                self, "确认退出",
                "后台确认处理仍在进行中，确定要退出吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return
        event.accept()

def main():
    """程序入口"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    app.setApplicationName("Markdown Frontmatter Processor")
    app.setApplicationVersion("1.4.0")
    app.setOrganizationName("Data Tools")
    
    # 支持命令行指定语言
    lang_code = 'zh'  # 默认中文
    if len(sys.argv) > 1 and sys.argv[1] in LanguageManager.LANGUAGES:
        lang_code = sys.argv[1]
    lang = LanguageManager(lang_code)
    
    window = MainWindow(lang)
    
    if len(sys.argv) > 2 and os.path.isdir(sys.argv[2]):
        window.input_dir_edit.setText(str(Path(sys.argv[2])))
        if len(sys.argv) > 3 and os.path.isdir(sys.argv[3]):
            window.output_dir_edit.setText(str(Path(sys.argv[3])))
    elif len(sys.argv) > 2 and os.path.isfile(sys.argv[2]) and sys.argv[2].endswith(('.yaml', '.yml')):
        window.config_edit.setText(sys.argv[2])
        window.load_config()
    
    window.show()
    
    if '--batch' in sys.argv:
        if sys.argv and window.input_dir_edit.text() and window.output_dir_edit.text():
            window.start_processing()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
