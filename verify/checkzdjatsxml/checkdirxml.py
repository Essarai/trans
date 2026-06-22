import os
import re
from lxml import etree


class MyErrorHandler(object):
    def __init__(self):
        self.errors = []

    def error(self, error):
        self.errors.append(error)


def replace_entities(content, entities):
    # 替换实体引用
    for entity, replacement in entities.items():
        pattern = r"&{};".format(entity)
        content = re.sub(pattern, replacement, content)
    return content


def validate_xml_with_xsd(xml_file, xmlschema, entities=None):
    # 创建一个XML Schema解析器

    # 创建一个XML文档解析器，忽略外部实体
    parser = etree.XMLParser(no_network=True, resolve_entities=False)

    try:
        # 读取XML文件内容
        with open(xml_file, 'r', encoding='utf-8') as file:
            content = file.read()

        # 如果有实体映射表，则替换实体引用
        if entities:
            content = replace_entities(content, entities)

        # 使用自定义解析器打开XML文件
        xml_doc = etree.fromstring(content.encode('utf-8'), parser=parser)
    except Exception as e:
        print(f"无法解析文件 {xml_file}: {e}")
        return False

    # 使用XML Schema验证XML文档
    error_handler = MyErrorHandler()
    try:
        valid = xmlschema.validate(xml_doc)
        if valid:
            #print(f"验证成功:{os.path.basename(xml_file)}")
            return True
        else:
            print(f"【{xml_file}】验证失败")
            for error in xmlschema.error_log:
                if 'Entity' in str(error):
                    pass
                elif 'xmlParseEntityRef' in str(error):
                    pass
                elif 'EntityRef' in str(error):
                    pass
                else:
                    print(f"  错误: {error}")
    except Exception as e:
        print(f"[{xml_file}] 验证失败: {e}")
        for error in error_handler.errors:
            if 'Entity' in str(error):
                pass
            elif 'xmlParseEntityRef' in str(error):
                pass
            elif 'EntityRef' in str(error):
                pass
            else:
                print(f"  错误: {error}")

        return False

    return True


def get_schema_from_xsdfile(xsd_file):
    with open(xsd_file, 'r') as schema_file:
        xmlschema_doc = etree.parse(schema_file)
    xmlschema = etree.XMLSchema(xmlschema_doc)

    return xmlschema
def process_directory(directory,xsd_file,entities=None):
    """
    遍历指定目录及其所有子目录下的所有XML文件，并打印出每个XML文件的路径名。

    参数:
    directory (str): 目录路径。
    """
    # 遍历指定目录及其所有子目录
    xmlschema = get_schema_from_xsdfile(xsd_file)
    for root, dirs, files in os.walk(directory):
        # 遍历当前目录下的所有文件
        for file in files:
            # 检查文件是否以 .xml 结尾
            if file.endswith('.xml') and not 'issue-' in file:
                # 使用 os.path.join() 确保路径的跨平台兼容性
                file_path = os.path.join(root, file)
                validate_xml_with_xsd(file_path,xmlschema,entities)

if __name__ == "__main__":

    xsd_file = r'verify/checkzdjatsxml/ZDjats/ZD_JATS-archivearticle1-3.xsd'#r'JATS-Archiving-1-3-MathML3-XSD/JATS-archivearticle1-3-mathml3.xsd'#r'ZDjats/ZD_JATS-archivearticle1-3.xsd'
    entities = {
        'copy': '©',
        'sect': '§',
        'Dagger': '†',
        'dagger': '†',
        'daggerdbl': '‡',
        'ellipsis': '…',
        'oacute': 'ó',
        'prime': '′',
        'aacute': 'á',
        'prime-double': '″',
        'prime-triple': '‴',
        'prime-quad': '⁗',
        'prime-quint': '⁗',
        'rsquo': '’',
        'lsquo': '‘',
        'ldquo': '“',
        'rdquo': '”',
        'uuml': 'ü',
        'auml': 'ä',
        'ouml': 'ö',
        'Auml': 'Ä',
        'Ouml': 'Ö',
        'Uuml': 'Ü',
        'omega': 'ω',
        'Omega': 'Ω',
        'alpha': 'α',
        'beta': 'β',
        'gamma': 'γ',
        'delta': 'δ',
        'epsilon': 'ε',
        'zeta': 'ζ',
        'eta': 'η',
        'theta': 'θ',
        'iota': 'ι',
        'kappa': 'κ',
        'lambda': 'λ',
        'mu': 'μ',
        'nu': 'ν',
        'xi': 'ξ',
        'omicron': 'ο',
        'pi': 'π',
        'rho': 'ρ',
        'sigma': 'σ',
        'ndash': '–',
        'reg': '®',
        'minus': '−',
        'deg': '°',
        'le': '≤',
        'ge': '≥',
        'approx': '≈',
        'times': '×',
        'divide': '÷',
        'prime-quad': '⁗',
        'plusmn': '±',
        'sup2': '²',
        'trade': '™',
        'copyright': '©',
        'amp': '&',
        'lt': '<',
        'gt': '>',
        'quot': '"',
        'apos': "'",
        'nbsp': ' ',
        'apos': "'",
        'quot': '"',
        'apos': "'",
        'quot': '"',
        'chi': 'χ',
        'phi': 'φ',
        'psi': 'ψ',
        'rho': 'ρ',
        'sigma': 'σ',
        'tau': 'τ',
        'upsilon': 'υ',
        'omega': 'ω',
        'theta': 'θ',
        'xi': 'ξ',
        'dash': '–',
        'iacute': 'í',
        'micro': 'µ',
        'Aacute': 'Á',
        'eacute': 'é',
        'iacute': 'í',
        'oacute': 'ó',
        'uacute': 'ú',
        'Uacute': 'Ú',
        'auml': 'ä',
        'ouml': 'ö',
        'uuml': 'ü',
        'Auml': 'Ä',
        'Ouml': 'Ö',
        'sum': '∑',
        'int': '∫',
        'infin': '∞',
        'part': '∂',
        'prime': '′',
        'prime-double': '″',
        'oslash': 'ø',
        'middot': '·',
        'tau': 'τ',
        'Oslash': 'Ø',
        'uarr': '↑',
        'bull': '•',
        'Delta': 'Δ',
        'darr': '↓',
        'mdash': '—',
        'rarr': '→',
        'larr': '←',
        'para': '¶',
        'deg': '°',
        'prime': '′',
        'prime-double': '″',
        'prime-triple': '‴',
        'prime-quad': '⁗',
        'prime-quint': '⁗',
        'prime-sext': '⁗',
        'prime-sept': '⁗',
        'prime-oct': '⁗',
        'prime-non': '⁗',
        'Psi': 'Ψ',
        'yacute': 'ý',
        'ograve': 'ò',
        'sup1': '¹',
        'egrave': 'è',
        'euml': 'ë',
        'harr': '↔',
        'yacute': 'ý',
        'Phi': 'Φ',
        'Scaron': 'Š',
        'scaron': 'š',
        'hellip': '…',
        'Scaron': 'Š',
        'ccedil': 'ç',
        'agrave': 'à',
        'szlig': 'ß',
        'Eacute': 'É',
        'ccedil': 'ç',
        'iuml': 'ï',
        'amp': '&',
        'rsquo': '’',
    }


    directory_to_process = directory = r'~/Desktop/trans/code/output'
    process_directory(directory_to_process, xsd_file,entities)