// 1. 자막을 표시할 레이어(Div) 생성
const subDiv = document.createElement('div');
subDiv.id = 'custom-subtitle-layer';
document.body.appendChild(subDiv);

// 2. 영상 요소 찾기
const video = document.querySelector('video');

if (video) {
  video.addEventListener('timeupdate', () => {
    const currentTime = video.currentTime;
    
    // 실제로는 여기서 자막 파일(JSON/SRT)의 시간과 비교해야 합니다.
    // 예시: 10초~12초 사이에 일본어 자막 출력
    if (currentTime >= 10 && currentTime <= 12) {
      subDiv.innerText = "こんにちは (곤니치와)";
      subDiv.style.display = "block";
    } else {
      subDiv.style.display = "none";
    }
  });
}