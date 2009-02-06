#ifndef RSA_H
#define RSA_H

#ifdef __cplusplus
extern "C" {
#endif

struct RSAKey {
  int bits;
  int bytes;
#ifdef MSCRYPTOAPI
  unsigned long exponent;
  unsigned char *modulus;
#else
  Bignum modulus;
  Bignum exponent;
  Bignum private_exponent;
  Bignum p;
  Bignum q;
  Bignum iqmp;
#endif
  char *comment;
};

extern LRESULT WINAPI AL_rsa_get_public_blob(struct RSAKey *key, char *blob, int *len);
extern LRESULT WINAPI AL_rsa_encrypt(int type, unsigned char *data, int *len, struct RSAKey *key);
extern LRESULT WINAPI AL_rsa_decrypt(int type, unsigned char *data, int *outlen, struct RSAKey *key);
extern LRESULT WINAPI AL_rsa_public_key_blob(struct RSAKey *key, unsigned char *pub_blob, int *len);
extern LRESULT WINAPI AL_rsa_private_key_blob(struct RSAKey *key, unsigned char *priv_blob, int *len);
extern LRESULT WINAPI AL_rsa_createkey(unsigned char *pub_blob, int pub_len, unsigned char *priv_blob, int priv_len, struct RSAKey *key);
extern LRESULT WINAPI AL_rsa_freekey(struct RSAKey *key);
extern LRESULT WINAPI AL_rsa_sign(struct RSAKey *key, char *data, int datalen, char *sig, int *siglen);
extern LRESULT WINAPI AL_rsa_verifysig(void *key, char *sig, int siglen, char *data, int datalen);

#ifdef __cplusplus
}
#endif

#endif