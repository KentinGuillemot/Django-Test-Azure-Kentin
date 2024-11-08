from django import template
from django.template.defaultfilters import stringfilter

register = template.Library()


@register.filter
def empty_if_zero(value):
    # Check if value is 0 or 0.0, and return an empty string
    if value == 0 or value == 0.0:
        return ''
    return value

@register.filter
def espace_separateur(value):
    """
    Format the number with a space as a thousands separator and retain decimal places.
    """
    if value is not None:
        value = str(value)
        for i in range(0, len(value)):
            if i%3 == 0 and i != 0:
                value = value[:i] + ' ' + value[i:]

    return value