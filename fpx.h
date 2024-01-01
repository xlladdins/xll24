// fpx.h - Excel FP12 data type.
#pragma once

#pragma warning(push)
// nonstandard extension used: zero-sized array in struct/union
#pragma warning(disable: 4200) 
struct fpx {
	int rows;
	int columns;
	double array[];
};
#pragma warning(pop)

// Row-major order
extern int fpx_index(struct fpx* fpx, int r, int c);
struct fpx* fpx_malloc(int r, int c);
struct fpx* fpx_realloc(struct fpx* fpx, int r, int c);
void fpx_free(struct fpx*);

