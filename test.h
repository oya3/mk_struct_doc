#ifndef _COMMON_H_
#define _COMMON_H_

typedef struct _Test00
{
	bool member00; // �����o�[�O�O
    unsigned char member01; /* �����o�[�O�P */
	short member02; // �����o�[�O�Q
    unsigned short member03; /* �����o�[�O�R */
	int member04; // �����o�[�O�S
} Test00;

typedef struct {
	bool member10; // �����o�[�P�O
    unsigned char member11; /* �����o�[�P�P */
	short member12; // �����o�[�P�Q
    unsigned short member13; /* �����o�[�P�R */
	int member14; // �����o�[�P�S
	Test00 member15; /* �����o�[�P�T */
} Test01;

#endif // _COMMON_H_
