#ifndef _COMMON_H_
#define _COMMON_H_


/**
 * @struct  Test00
 * @brief   テスト００のデータ構造
 * @details この構造は最大長を考慮して定義してある
 */
typedef struct _Test00
{
	bool member00; // メンバー００
    unsigned char member01; /* メンバー０１ */
	short member02; // メンバー０２
    unsigned short member03; /* メンバー０３ */
	int member04; // メンバー０４
} Test00,TEST00;

/**
 * @struct  Test01
 * @brief   テスト０１のデータ構造
 * @details １行目
 *          ２行目
 *          ３行目
 */
typedef struct {
	bool member10; // メンバー１０
    unsigned char member11; /* メンバー１１ */
	short member12; // メンバー１２
    unsigned short member13; /* メンバー１３ */
	int member14; // メンバー１４
	Test00 member15; /* メンバー１５ */
} Test01;

#endif // _COMMON_H_
