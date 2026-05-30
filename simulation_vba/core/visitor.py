


class visitor(object):
    """
    The class template for visitor objects for the visitor design pattern.
    Visitors can be accepted by the accept method of VBA_Object objects.
    """

    def visit(self):
        raise NotImplementedError("Not implemented.")
