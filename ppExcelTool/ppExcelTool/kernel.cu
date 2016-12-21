#include "E:\Dropbox\GitHub\myHelper\cuda8Helper\cuda8Helper\myCuda.cu"

int main() {
  myCuda::gpuInfo();
  //unified memory
  const int N = 10;
  Dim(x, float, N);
  myCuda::print::print_float<<<1,N>>>(x, N);
  


  cudaFree(x);
  system("pause");
  return 0;
}
