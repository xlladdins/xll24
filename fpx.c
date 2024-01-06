// fpx.c - C VLA implementation of the fpx.h interface.
#include <stdlib.h>
#include "fpx.h"

int fpx_index(struct fpx* p, int i, int j)
{
	return p->columns * i + j;
}

struct fpx* fpx_malloc(int r, int c)
{
	struct fpx* fpx = malloc(sizeof(struct fpx) + r * c * sizeof(double));

	if (fpx) {
		fpx->rows = r;
		fpx->columns = c;
	}

	return fpx;
}

struct fpx* fpx_realloc(struct fpx* p, int r, int c)
{
	struct fpx* _p = realloc(p, sizeof(struct fpx) + r * c * sizeof(double));

	if (_p) {
		_p->rows = r;
		_p->columns = c;
	}

	return _p ? _p : (void*)0;
}

void fpx_free(struct fpx* p)
{
	free(p);
}
