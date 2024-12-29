import re
ticket_num_reg = re.compile(
    r'''[I|i][D|d][0-9]{9}|[I|i][B|b][0-9]{9}|[R|r][P|p][0-9]{9}''', re.MULTILINE)
subject_regex = re.compile(r''''[#ID[0-9]+]''', re.MULTILINE)
idriveinc_domain_regex = re.compile(
    r'''\b[A-Za-z0-9._%+-]+@(idrive\.com|idriveinc\.com)\b''', re.MULTILINE)


print(re.findall(idriveinc_domain_regex, "   sandeep.kumar@idriveinc.com "))
